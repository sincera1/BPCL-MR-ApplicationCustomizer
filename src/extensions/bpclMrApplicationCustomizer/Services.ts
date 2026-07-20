import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IMenuItem {
    Id: number;
    LinkTitle: string;
    TitleHindi: string;
    TitleMarathi: string;
    RedirectURL: {
        Url: string;
        Description: string;
    };
    Sequence: number;
    IsActive: boolean;
}

export default class Services {

    private _sp: SPFI;

    constructor(sp: SPFI) {
        this._sp = sp;
    }

    public async getMasterMenu(): Promise<IMenuItem[]> {

        const items: IMenuItem[] = await this._sp.web.lists
            .getByTitle("MR_SL_MasterMenu")
            .items
            .select(
                "Id",
                "LinkTitle",
                "TitleHindi",
                "TitleMarathi",
                "RedirectURL",
                "Sequence",
                "IsActive"
            )
            .filter("IsActive eq 1")
            .orderBy("Sequence", true)();

        return items;
    }
}