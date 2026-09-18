import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/attachments";
import "@pnp/sp/site-users/web";
import { IListAnnouncement, IListAnnouncementAttachment, IListAnnouncementAudience } from "../models/IListAnnouncement";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface IListAnnouncementListItem {
  Id: number;
  Title: string;
  Body?: string;
  HighlightType: string;
  Site: string;
  Priority: string;
  Status: string;
  RejectionReason?: string;
  TargetAudienceType?: string;
  TargetAudience?: IListAnnouncementAudience[];
  BannerImageUrl?: { Url: string; Description?: string };
  Author?: { Id: number; Title: string };
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
  Created: string;
}

interface IListAnnouncementItemData {
  Title: string;
  Body: string;
  HighlightType: string;
  Site: string;
  Priority: string;
  Status: string;
  RejectionReason?: string;
  TargetAudienceType?: string;
  TargetAudienceId?: number[];
  BannerImageUrl?: { Url: string; Description: string } | undefined;
}


export class ListAnnouncementService {
  private listName: string;
  private sp: SPFI;

  constructor(context: WebPartContext, listName: string) {
    this.listName = listName;
    this.sp = spfi().using(SPFx(context));
  }

  public async getListAnnouncements(): Promise<IListAnnouncement[]> {
    let currentUserTitle = "";
    try {
      const currentUser = await this.sp.web.currentUser.select("Title")();
      currentUserTitle = currentUser.Title?.toLowerCase().trim() ?? "";
    } catch {
      console.warn("Could not fetch current user title");
    }

    const items: IListAnnouncementListItem[] = await this.sp.web.lists
      .getByTitle(this.listName)
      .items.select(
        "Id", "Title", "Body", "HighlightType", "Site",
        "Priority", "Status", "RejectionReason",
        "TargetAudienceType", "TargetAudience/Id", "TargetAudience/Title",
        "BannerImageUrl", "Author/Id", "Author/Title",
        "AttachmentFiles", "Created"
      )
      .expand("TargetAudience", "Author", "AttachmentFiles")();

    const announcements: IListAnnouncement[] = items.map((i) => ({
      Id: i.Id,
      Title: i.Title,
      Body: i.Body,
      HighlightType: i.HighlightType,
      Site: i.Site,
      Priority: (i.Priority === "Pinned" ? "Pinned" : "Not Pinned") as "Pinned" | "Not Pinned",
      Status: i.Status as "Draft" | "Published" | "Rejected",
      RejectionReason: i.RejectionReason || undefined,
      TargetAudienceType: (i.TargetAudienceType || "All") as "All" | "Specific" | "Except",
      TargetAudience: i.TargetAudience ?? [],
      BannerImageUrl: i.BannerImageUrl?.Url,
      Author: i.Author ? { Id: i.Author.Id, Title: i.Author.Title } : undefined,
      Attachments: (i.AttachmentFiles ?? []).map((f) => ({
        FileName: f.FileName,
        ServerRelativeUrl: f.ServerRelativeUrl,
      } as IListAnnouncementAttachment)),
      Created: new Date(i.Created),
    }));

    const filtered = announcements.filter((h) => {
      const type = h.TargetAudienceType || "All";
      if (type === "All") return true;
      const audienceTitles = (h.TargetAudience || []).map((p) => p.Title?.toLowerCase().trim());
      const isInAudience = audienceTitles.includes(currentUserTitle);
      if (type === "Specific") return isInAudience;
      if (type === "Except") return !isInAudience;
      return true;
    });

    filtered.sort((a, b) => {
      const aPinned = a.Priority === "Pinned" ? 0 : 1;
      const bPinned = b.Priority === "Pinned" ? 0 : 1;
      if (aPinned !== bPinned) return aPinned - bPinned;
      return (b.Created?.getTime() ?? 0) - (a.Created?.getTime() ?? 0);
    });

    return filtered;
  }

  public async addListAnnouncement(
    announcement: Partial<IListAnnouncement>,
    bannerFile?: File,
    attachments?: File[]
  ): Promise<void> {
    let bannerUrl: string | undefined;
    if (bannerFile) {
      try { bannerUrl = await this.uploadBanner(bannerFile); }
      catch (error) { console.error("Banner upload failed:", error); }
    }

    const itemData: IListAnnouncementItemData = {
      Title: announcement.Title ?? "",
      Body: announcement.Body || "",
      HighlightType: announcement.HighlightType ?? "",
      Site: announcement.Site ?? "",
      Priority: announcement.Priority ?? "Medium",
      Status: announcement.Status ?? "Draft",
      RejectionReason: announcement.RejectionReason,
      TargetAudienceType: announcement.TargetAudienceType || "All",
    };

    if (announcement.TargetAudience && announcement.TargetAudience.length > 0) {
      const audienceIds = announcement.TargetAudience
        .filter((p) => p && typeof p.Id === "number" && p.Id > 0)
        .map((p) => p.Id);
      if (audienceIds.length > 0) itemData.TargetAudienceId = audienceIds;
    }

    if (bannerUrl) {
      itemData.BannerImageUrl = { Url: bannerUrl, Description: announcement.Title || "Announcement Banner" };
    }

    const addResult = await this.sp.web.lists.getByTitle(this.listName).items.add(itemData);

    if (attachments && attachments.length > 0) {
      await this.uploadAttachments(addResult.Id, attachments);
    }
  }

  public async updateListAnnouncement(
    announcementId: number,
    announcement: Partial<IListAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> {
    let bannerUrl = announcement.BannerImageUrl;
    if (bannerFile) {
      try { bannerUrl = await this.uploadBanner(bannerFile); }
      catch (error) { console.error("Banner upload failed:", error); }
    }

    const itemData: IListAnnouncementItemData = {
      Title: announcement.Title ?? "",
      Body: announcement.Body || "",
      HighlightType: announcement.HighlightType ?? "",
      Site: announcement.Site ?? "",
      Priority: announcement.Priority ?? "Medium",
      Status: announcement.Status ?? "Draft",
      RejectionReason: announcement.RejectionReason,
      TargetAudienceType: announcement.TargetAudienceType || "All",
    };

    if (announcement.TargetAudience && announcement.TargetAudience.length > 0) {
      const audienceIds = announcement.TargetAudience
        .filter((p) => p && typeof p.Id === "number" && p.Id > 0)
        .map((p) => p.Id);
      if (audienceIds.length > 0) itemData.TargetAudienceId = audienceIds;
    }

    if (bannerUrl) {
      itemData.BannerImageUrl = { Url: bannerUrl, Description: announcement.Title || "Announcement Banner" };
    } else if (!bannerFile && !announcement.BannerImageUrl) {
      itemData.BannerImageUrl = undefined;
    }

    await this.sp.web.lists.getByTitle(this.listName).items.getById(announcementId).update(itemData);

    if (deletedAttachmentNames && deletedAttachmentNames.length > 0) {
      await this.deleteAttachments(announcementId, deletedAttachmentNames);
    }

    if (attachments && attachments.length > 0) {
      await this.uploadAttachments(announcementId, attachments);
    }
  }

  public async approveAnnouncement(announcementId: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.listName).items.getById(announcementId).update({ Status: "Published", RejectionReason: "" });
  }

  public async rejectAnnouncement(announcementId: number, reason: string): Promise<void> {
    await this.sp.web.lists.getByTitle(this.listName).items.getById(announcementId).update({ Status: "Rejected", RejectionReason: reason });
  }

  public async uploadAttachments(itemId: number, files: File[]): Promise<void> {
    for (const file of files) {
      try {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .attachmentFiles.add(file.name, file);
      } catch (error) {
        console.error("Failed to upload attachment " + file.name + ":", error);
      }
    }
  }

  public async deleteAttachments(itemId: number, fileNames: string[]): Promise<void> {
    for (const fileName of fileNames) {
      try {
        await this.sp.web.lists
          .getByTitle(this.listName)
          .items.getById(itemId)
          .attachmentFiles.getByName(fileName)
          .delete();
      } catch (error) {
        console.error("Failed to delete attachment " + fileName + ":", error);
      }
    }
  }

  private async uploadBanner(file: File): Promise<string> {
    const folderPath = "SiteAssets/ListAnnouncementImages";
    const fileName = Date.now().toString() + "_" + file.name;
    try {
      try {
        await this.sp.web.getFolderByServerRelativePath(folderPath).select("Exists")();
      } catch {
        await this.sp.web.folders.addUsingPath(folderPath);
      }
      const uploadResult = await this.sp.web
        .getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(fileName, file, { Overwrite: true });
      return uploadResult.ServerRelativeUrl;
    } catch (error) {
      console.error("Banner upload error:", error);
      throw new Error("Failed to upload banner image");
    }
  }
}
