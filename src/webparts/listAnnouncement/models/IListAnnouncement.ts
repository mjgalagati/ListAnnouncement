export interface IListAnnouncementAuthor {
  Id: number;
  Title: string;
}

export interface IListAnnouncementAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IListAnnouncementAudience {
  Id: number;
  Title: string;
}

export interface IListAnnouncement {
  Id: number;
  Title: string;
  Body?: string;
  HighlightType: string;
  Site: string;
  Priority: "Pinned" | "Not Pinned";
  Status: "Draft" | "Published";
  TargetAudienceType?: "All" | "Specific" | "Except";
  TargetAudience?: IListAnnouncementAudience[];
  BannerImageUrl?: string;
  Attachments?: IListAnnouncementAttachment[];
  Author?: IListAnnouncementAuthor;
  Created?: Date;
}
