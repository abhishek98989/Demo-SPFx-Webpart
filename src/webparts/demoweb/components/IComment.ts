// src/webparts/listItemComments/models/IComment.ts
export interface IComment {
    id: string;
    text: string;
    author: {
      id: any;
      name: string;
      email: string;
    };
    createdDate: Date;
    likes: string[]; // User IDs who liked the comment
    replies?: IReply[];
  }
  
  export interface IReply {
    id: string;
    text: string;
    author: {
      id: any;
      name: string;
      email: string;
    };
    createdDate: Date;
    likes: string[]; // User IDs who liked the reply
  }