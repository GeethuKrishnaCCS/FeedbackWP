import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups/web";


export class FeedbackWebpartService extends BaseService {
    private _spfi: SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this._spfi = getSP(context);
    }


    public addListItem(data: any, listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    }
    public getListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }
    public updateItem(listname: string, data: any, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public async getCurrentUser(): Promise<any> {
        return this._spfi.web.currentUser();
    }

    public async getSiteUsers(): Promise<any> {
        return this._spfi.web.siteUsers();
    }

    public async getSiteGroups(): Promise<any> {
        return this._spfi.web.siteGroups();
    }

    public async isUserOwnerOfGroup(): Promise<boolean> {
        return  this._spfi.web.currentUser.groups();
        // const userGroups = await this._spfi.web.currentUser.groups();
        // return userGroups.some(group => group.Title === groupName);
    }

}