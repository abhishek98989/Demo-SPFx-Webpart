import * as React from 'react';
import styles from './Demoweb.module.scss';
import { IListItemCommentsProps } from './IDemowebProps';
import { CommentsService } from './CommentsService';
import { IComment, IReply } from './IComment';
import { 
  PrimaryButton, 
  Spinner, 
  SpinnerSize, 
  Stack, 
  TextField, 
  PersonaSize, 
  Persona, 
  IconButton,
  Text,
  ActionButton,
  IPersonaProps,
  IIconProps
} from '@fluentui/react';
import { useEffect, useState } from 'react';
import { AIEnabledInputField } from './AIEnabledInputField';
import { Newtable } from './Newtable';

// Custom Comment Card component to replace the PnP one
interface ICustomCommentCardProps {
  text: string;
  authorName: string;
  authorEmail: string;
  createdDate: any;
  likeCount: number;
  isLiked: boolean;
  onLike: () => void;
  onReply?: () => void;
  isReply?: boolean;
}

const CustomCommentCard: React.FC<ICustomCommentCardProps> = (props) => {
  const { 
    text, 
    authorName, 
    authorEmail, 
    createdDate, 
    likeCount, 
    isLiked, 
    onLike, 
    onReply, 
    isReply = false 
  } = props;

  const likeIcon: IIconProps = { iconName: isLiked ? 'LikeSolid' : 'Like' };
  const replyIcon: IIconProps = { iconName: 'Reply' };

  const personaProps: IPersonaProps = {
    text: authorName,
    secondaryText: new Date(createdDate).toLocaleString(),
    size: PersonaSize.size32,
    imageInitials: authorName.substring(0, 2),
    showSecondaryText: true,
  };

  return (
    <div className={isReply ? styles.replyCard : styles.commentCard}>
      <div className={styles.commentHeader}>
        <Persona {...personaProps} />
      </div>
      <div className={styles.commentContent}>
        <Text>{text}</Text>
      </div>
      <div className={styles.commentActions}>
        <ActionButton 
          iconProps={likeIcon} 
          onClick={onLike}
          text={`${likeCount} ${likeCount === 1 ? 'Like' : 'Likes'}`}
        />
        {!isReply && onReply && (
          <ActionButton
            iconProps={replyIcon}
            onClick={onReply}
            text="Reply"
          />
        )}
      </div>
    </div>
  );
};

export default function ListItemComments(props: IListItemCommentsProps): React.ReactElement<IListItemCommentsProps> {
  const { context, listName, itemId } = props;

  const [comments, setComments]:any = useState<IComment[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [newCommentText, setNewCommentText] = useState<string>('');
  const [addingComment, setAddingComment] = useState<boolean>(false);
  const [replyMap, setReplyMap] = useState<Record<string, string>>({});

  const commentsService = new CommentsService(context);

  useEffect(() => {
    if (!listName || !itemId) {
      setError('Please configure the webpart with a list name and ensure an item ID is provided in the URL.');
      setLoading(false);
      return;
    }

    loadComments();
  }, [listName, itemId]);

  const loadComments = async (): Promise<void> => {
    try {
      setLoading(true);
      const loadedComments = await commentsService.getComments(listName, itemId);
      setComments(loadedComments);
      setError(null);
    } catch (err) {
      console.error('Error loading comments:', err);
      setError('Failed to load comments. Please refresh the page or check your configuration.');
    } finally {
      setLoading(false);
    }
  };

  const handleAddComment = async (): Promise<void> => {
    if (!newCommentText.trim()) return;

    try {
      setAddingComment(true);
      const newComment = await commentsService.addComment(listName, itemId, newCommentText);
      setComments([...comments, newComment]);
      setNewCommentText('');
    } catch (err) {
      console.error('Error adding comment:', err);
      setError('Failed to add comment. Please try again.');
    } finally {
      setAddingComment(false);
    }
  };

  const handleAddReply = async (commentId: string): Promise<void> => {
    const replyText = replyMap[commentId];
    if (!replyText || !replyText.trim()) return;

    try {
      const newReply = await commentsService.addReply(listName, itemId, commentId, replyText);
      
      // Update the comments array with the new reply
      const updatedComments = comments.map((comment: any) => {
        if (comment.id === commentId) {
          return {
            ...comment,
            replies: [...(comment.replies || []), newReply]
          };
        }
        return comment;
      });
      
      setComments(updatedComments);
      
      // Clear the reply text for this comment
      const updatedReplyMap = { ...replyMap };
      delete updatedReplyMap[commentId];
      setReplyMap(updatedReplyMap);
    } catch (err) {
      console.error('Error adding reply:', err);
      setError('Failed to add reply. Please try again.');
    }
  };

  const handleToggleLike = async (commentId: string, isReply: boolean = false, parentId?: string): Promise<void> => {
    try {
      if (isReply && parentId) {
        // Find the reply
        const comment = comments?.find((c: any) => c.id === parentId);
        const reply = comment?.replies?.find((r: any) => r.id === commentId);
        if (!reply) return;

        // Check if current user has liked this reply
        const currentUserId = context.pageContext.user.loginName;
        const isLiked = reply.likes.includes(currentUserId);

        // Toggle like on the server
        await commentsService.toggleReplyLike(listName, itemId, parentId, commentId, isLiked);

        // Update state
        const updatedComments = comments.map((c: any) => {
          if (c.id === parentId) {
            return {
              ...c,
              replies: c.replies?.map((r: any) => {
                if (r.id === commentId) {
                  const updatedLikes = isLiked
                    ? r.likes.filter((id: any) => id !== currentUserId)
                    : [...r.likes, currentUserId];
                  return { ...r, likes: updatedLikes };
                }
                return r;
              })
            };
          }
          return c;
        });
        
        setComments(updatedComments);
      } else {
        // Find the comment
        const comment = comments.find((c: any) => c.id === commentId);
        if (!comment) return;

        // Check if current user has liked this comment
        const currentUserId = context.pageContext.user.loginName;
        const isLiked = comment.likes.includes(currentUserId);

        // Toggle like on the server
        await commentsService.toggleCommentLike(listName, itemId, commentId, isLiked);

        // Update state
        const updatedComments = comments.map((c: any) => {
          if (c.id === commentId) {
            const updatedLikes = isLiked
              ? c.likes.filter((id: any) => id !== currentUserId)
              : [...c.likes, currentUserId];
            return { ...c, likes: updatedLikes };
          }
          return c;
        });
        
        setComments(updatedComments);
      }
    } catch (err) {
      console.error('Error toggling like:', err);
      setError('Failed to update like. Please try again.');
    }
  };

  const renderContent = (): JSX.Element => {
    if (loading) {
      return <Spinner size={SpinnerSize.large} label="Loading comments..." />;
    }

    if (error) {
      return <div className={styles.error}>{error}</div>;
    }

    if (!listName || !itemId) {
      return <div>Please configure the webpart with a list name and ensure an item ID is provided in the URL.</div>;
    }

    return (
      <div className={styles.commentsContainer}>
        <h2 className={styles.commentsHeader}>Comments</h2>
        
        {/* Add Comment Section */}
        <div className={styles.addCommentSection}>
        <AIEnabledInputField
        label="Content"
        value={newCommentText}
        onChange={(_, newValue) => setNewCommentText(newValue || '')}
        multiline
        rows={8}
        perplexityApiKey={''}
        required
      />
          {/* <TextField
            multiline
            rows={3}
            placeholder="Add a comment..."
            value={newCommentText}
            onChange={(_, newValue) => setNewCommentText(newValue || '')}
            disabled={addingComment}
          /> */}
          <Stack horizontal horizontalAlign="end" className={styles.addCommentButtonContainer}>
            <PrimaryButton 
              text="Add Comment" 
              onClick={handleAddComment} 
              disabled={!newCommentText.trim() || addingComment}
            />
          </Stack>
        </div>

        {/* Comments List */}
        <div className={styles.commentsList}>
          {comments.length === 0 ? (
            <div className={styles.noComments}>No comments yet. Be the first to comment!</div>
          ) : (
            comments.map((comment: any) => (
              <div key={comment.id} className={styles.commentCardWrapper}>
                <CustomCommentCard
                  text={comment.text}
                  authorName={comment.author.name}
                  authorEmail={comment.author.email}
                  createdDate={comment.createdDate}
                  likeCount={comment.likes.length}
                  isLiked={comment.likes.includes(context.pageContext.user.loginName)}
                  onLike={() => handleToggleLike(comment.id)}
                  onReply={() => {
                    // Set an empty reply for this comment if one doesn't exist
                    if (!replyMap[comment.id]) {
                      setReplyMap({ ...replyMap, [comment.id]: '' });
                    }
                  }}
                />
                
                {/* Reply input field (if in reply mode for this comment) */}
                {replyMap[comment.id] !== undefined && (
                  <div className={styles.replyInputContainer}>
                    <TextField
                      placeholder="Write a reply..."
                      value={replyMap[comment.id]}
                      onChange={(_, newValue) => setReplyMap({ ...replyMap, [comment.id]: newValue || '' })}
                      className={styles.replyInput}
                    />
                    <Stack horizontal className={styles.replyButtonContainer}>
                      <PrimaryButton
                        text="Reply"
                        onClick={() => handleAddReply(comment.id)}
                        disabled={!replyMap[comment.id]?.trim()}
                      />
                      <PrimaryButton
                        text="Cancel"
                        onClick={() => {
                          const updatedReplyMap = { ...replyMap };
                          delete updatedReplyMap[comment.id];
                          setReplyMap(updatedReplyMap);
                        }}
                      />
                    </Stack>
                  </div>
                )}
                
                {/* Replies list */}
                {comment.replies && comment.replies.length > 0 && (
                  <div className={styles.repliesContainer}>
                    {comment.replies.map((reply: any) => (
                      <div key={reply.id} className={styles.replyCardWrapper}>
                        <CustomCommentCard
                          text={reply.text}
                          authorName={reply.author.name}
                          authorEmail={reply.author.email}
                          createdDate={reply.createdDate}
                          likeCount={reply.likes.length}
                          isLiked={reply.likes.includes(context.pageContext.user.loginName)}
                          onLike={() => handleToggleLike(reply.id, true, comment.id)}
                          isReply={true}
                        />
                      </div>
                    ))}
                  </div>
                )}
              </div>
            ))
          )}
        </div>
      </div>
    );
  };

  return (
   <Newtable Context={context} />
  );
}