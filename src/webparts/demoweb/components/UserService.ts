// src/webparts/listItemComments/services/UserService.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { IUserInfo } from "./IUserInfo";

export class UserService {
  private sp: SPFI;

  constructor(private context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Get users from the current site
   */
  public async getSiteUsers(): Promise<IUserInfo[]> {
    try {
      const users = await this.sp.web.siteUsers();
      
      // Filter out system accounts and map to our user model
      return users
        .filter((user:any) => 
          user.UserPrincipalName && 
          !user.UserPrincipalName.includes('spoapp') && 
          !user.UserPrincipalName.includes('spo-grid') &&
          !user.LoginName.includes('_spvisitor') &&
          !user.LoginName.includes('_spwriter') &&
          user.Email !== ''
        )
        .map(user => ({
          id: user.Id,
          email: user.Email,
          loginName: user.LoginName,
          displayName: user.Title
        }));
    } catch (error) {
      console.error("Error getting site users:", error);
      return [];
    }
  }

  /**
   * Search for users by query
   * @param query The search query
   */
  public async searchUsers(query: string): Promise<IUserInfo[]> {
    try {
      // In a real implementation, you might want to use SharePoint's People Picker API
      // or Microsoft Graph to search for users
      // For now, we'll get all site users and filter them on the client
      const allUsers = await this.getSiteUsers();
      
      if (!query) return allUsers;
      
      const lowerQuery = query.toLowerCase();
      
      return allUsers.filter((user:any) => 
        user.displayName.toLowerCase().includes(lowerQuery) ||
        user.email.toLowerCase().includes(lowerQuery)
      );
    } catch (error) {
      console.error("Error searching users:", error);
      return [];
    }
  }
}