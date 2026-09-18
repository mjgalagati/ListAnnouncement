import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IComment {
  Id: number;
  CommentText: string;
  AuthorName: string;
  AuthorLoginName: string;
  Created: Date;
  isOwn: boolean;
  ParentCommentId: number | undefined;
}

export class CommentService {
  private _sp: ReturnType<typeof spfi>;
  private _listName: string;
  private _userEmail: string;
  private _alias: string;

  constructor(context: WebPartContext, listName: string, alias: string) {
    this._sp = spfi().using(SPFx(context));
    this._listName = listName;
    this._userEmail = context.pageContext.user.email;
    this._alias = alias.trim().toLowerCase();
  }

  async getComments(messageId: number): Promise<IComment[]> {
    // Fetch all items then filter client-side to avoid OData internal-name issues
    // Title scheme (built-in field, no internal-name ambiguity):
    //   top-level → "Comment on {alias} {messageId}"
    //   reply     → "Comment on {alias} {messageId}|P:{parentCommentId}"
    // Alias namespaces rows so multiple web parts can share this list.
    const titleBase = `Comment on ${this._alias} ${messageId}`;
    const items = await this._sp.web.lists
      .getByTitle(this._listName)
      .items
      .select("Id", "Title", "CommentText", "Created", "Author/Title", "Author/EMail")
      .expand("Author")
      .orderBy("Created", true)();

    return items
      .filter(item =>
        typeof item.Title === "string" &&
        (item.Title === titleBase || item.Title.startsWith(titleBase + "|P:"))
      )
      .map(item => {
        const parentMatch = typeof item.Title === "string" && item.Title.includes("|P:")
          ? Number(item.Title.split("|P:")[1])
          : undefined;
        return {
          Id: item.Id,
          CommentText: item.CommentText ?? "",
          AuthorName: item.Author?.Title ?? "Unknown",
          AuthorLoginName: item.Author?.EMail ?? "",
          Created: new Date(item.Created),
          isOwn: item.Author?.EMail?.toLowerCase() === this._userEmail.toLowerCase(),
          ParentCommentId: Number.isFinite(parentMatch) ? parentMatch : undefined,
        };
      });
  }

  async addComment(messageId: number, text: string, authorName: string, parentCommentId?: number): Promise<IComment> {
    const list = this._sp.web.lists.getByTitle(this._listName);
    const titleBase = `Comment on ${this._alias} ${messageId}`;
    const title = parentCommentId !== undefined
      ? `${titleBase}|P:${parentCommentId}`
      : titleBase;
    const payload: Record<string, unknown> = {
      Title: title,
      CommentText: text,
    };

    const added = await list.items.add(payload);

    return {
      Id: added.Id,
      CommentText: text,
      AuthorName: authorName,
      AuthorLoginName: this._userEmail,
      Created: new Date(),
      isOwn: true,
      ParentCommentId: parentCommentId,
    };
  }

  async deleteComment(commentId: number, replyIds: number[] = []): Promise<void> {
    const list = this._sp.web.lists.getByTitle(this._listName);
    await Promise.all(
      [commentId, ...replyIds].map(id => list.items.getById(id).delete())
    );
  }
}
