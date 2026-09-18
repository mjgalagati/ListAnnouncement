import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export type ReactionType = "Like" | "Heart" | "Wow";

export interface IReactionSummary {
  counts: Record<ReactionType, number>;
  myReaction: ReactionType | undefined;
  myItemId: number | undefined;
}

export class ReactionService {
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

  async getReactions(messageId: number): Promise<IReactionSummary> {
    // Filter by Title prefix ("{alias}_{messageId}_") — Title is a guaranteed built-in
    // field, avoiding the internal-name ambiguity that custom Number columns can have.
    // Alias leads the prefix so web parts sharing this list can't collide on messageId.
    const items = await this._sp.web.lists
      .getByTitle(this._listName)
      .items
      .select("Id", "Title", "ReactionType", "Author/EMail")
      .expand("Author")();

    const prefix = `${this._alias}_${messageId}_`;
    const counts: Record<ReactionType, number> = { Like: 0, Heart: 0, Wow: 0 };
    let myReaction: ReactionType | undefined = undefined;
    let myItemId: number | undefined = undefined;

    for (const item of items) {
      if (typeof item.Title !== "string" || !item.Title.startsWith(prefix)) continue;
      const type = item.ReactionType as ReactionType;
      if (counts[type] !== undefined) counts[type]++;
      if (item.Author?.EMail?.toLowerCase() === this._userEmail.toLowerCase()) {
        myReaction = type;
        myItemId = item.Id;
      }
    }

    return { counts, myReaction, myItemId };
  }

  async setReaction(messageId: number, type: ReactionType, current: IReactionSummary): Promise<IReactionSummary> {
    const list = this._sp.web.lists.getByTitle(this._listName);
    if (current.myItemId !== undefined) {
      await list.items.getById(current.myItemId).update({ ReactionType: type });
      return { ...current, myReaction: type };
    } else {
      const added = await list.items.add({
        Title: `${this._alias}_${messageId}_${this._userEmail}`,
        MessageId: messageId,
        ReactionType: type,
      });
      return { ...current, myReaction: type, myItemId: added.Id };
    }
  }

  async removeReaction(current: IReactionSummary): Promise<void> {
    if (current.myItemId === undefined) return;
    await this._sp.web.lists
      .getByTitle(this._listName)
      .items.getById(current.myItemId)
      .delete();
  }
}
