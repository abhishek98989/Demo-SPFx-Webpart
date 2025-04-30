// src/webparts/listItemComments/services/CommentsService.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { ICommentInfo } from "@pnp/sp/comments";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/comments";
import { IComment, IReply } from "./IComment";

export class CommentsService {
  private sp: SPFI;

  constructor(private context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Get comments for a list item
   * @param listName The list name
   * @param itemId The item ID
   */
  public async getComments(listName: string, itemId: any): Promise<IComment[]> {
    try {
      // Get the comments from the SharePoint item
      const comments = await this.sp.web.lists.getByTitle(listName).items.getById(itemId).comments();
      
      // Transform the comments to our model
      return await Promise.all(comments.map(async (comment:any) => {
        // Get replies for this comment
        const replies = await this.sp.web.lists.getByTitle(listName).items.getById(itemId).comments.getById(comment.id).replies();
        
        // Map replies to our model
        const formattedReplies = replies.map((reply:any) => {
          return {
            id: reply.id,
            text: reply.text,
            author: {
              id: reply.author.id,
              name: reply.author.name,
              email: reply.author.email
            },
            createdDate: new Date(reply.createdDate),
            likes: reply.likedBy ? reply.likedBy.map((user:any) => user.id) : []
          } as IReply;
        });

        // Return the comment with its replies
        return {
          id: comment.id,
          text: comment.text,
          author: {
            id: comment.author.id,
            name: comment.author.name,
            email: comment.author.email
          },
          createdDate: new Date(comment.createdDate),
          likes: comment.likedBy ? comment.likedBy.map((user:any) => user.id) : [],
          replies: formattedReplies
        } as IComment;
      }));
    } catch (error) {
      console.error("Error getting comments:", error);
      throw error;
    }
  }

  /**
   * Add a comment to a list item
   * @param listName The list name
   * @param itemId The item ID
   * @param commentText The comment text
   */
  public async addComment(listName: string, itemId: any, commentText: string): Promise<IComment> {
    try {
      const result = await this.sp.web.lists.getByTitle(listName).items.getById(itemId).comments.add(commentText);
      
      // Return the newly created comment
      return {
        id: result.id,
        text: result.text,
        author: {
          id: result.author.id,
          name: result.author.name,
          email: result.author.email
        },
        createdDate: new Date(result.createdDate),
        likes: [],
        replies: []
      };
    } catch (error) {
      console.error("Error adding comment:", error);
      throw error;
    }
  }

  /**
   * Add a reply to a comment
   * @param listName The list name
   * @param itemId The item ID
   * @param commentId The parent comment ID
   * @param replyText The reply text
   */
  public async addReply(listName: string, itemId: any, commentId: string, replyText: string): Promise<IReply> {
    try {
      const result = await this.sp.web.lists.getByTitle(listName).items.getById(itemId).comments.getById(commentId).replies.add(replyText);
      
      // Return the newly created reply
      return {
        id: result.id,
        text: result.text,
        author: {
          id: result.author.id,
          name: result.author.name,
          email: result.author.email
        },
        createdDate: new Date(result.createdDate),
        likes: []
      };
    } catch (error) {
      console.error("Error adding reply:", error);
      throw error;
    }
  }

  /**
   * Like or unlike a comment
   * @param listName The list name
   * @param itemId The item ID
   * @param commentId The comment ID
   * @param isLiked Whether the comment is already liked
   */
  public async toggleCommentLike(listName: string, itemId: any, commentId: string, isLiked: boolean): Promise<void> {
    try {
      const comment = this.sp.web.lists.getByTitle(listName).items.getById(itemId).comments.getById(commentId);
      
      if (isLiked) {
        await comment.unlike();
      } else {
        await comment.like();
      }
    } catch (error) {
      console.error("Error toggling comment like:", error);
      throw error;
    }
  }

  /**
   * Like or unlike a reply
   * @param listName The list name
   * @param itemId The item ID
   * @param commentId The parent comment ID
   * @param replyId The reply ID
   * @param isLiked Whether the reply is already liked
   */
  public async toggleReplyLike(listName: string, itemId: any, commentId: string, replyId: string, isLiked: boolean): Promise<void> {
   
  }
}