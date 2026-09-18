import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, Icon } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListAnnouncement } from '../models/IListAnnouncement';
import styles from './ListAnnouncementDetailsPanel.module.scss';
import { ReactionService, ReactionType, IReactionSummary } from '../services/ReactionService';
import { CommentService, IComment } from '../services/CommentService';

interface ListAnnouncementDetailsPanelProps {
  announcement?: IListAnnouncement;
  isOpen: boolean;
  onDismiss: () => void;
  onEdit: (announcement: IListAnnouncement) => void;
  currentUserId: number;
  isEditor: boolean;
  isApprover: boolean;
  hasApproval: boolean;
  context: WebPartContext;
  reactionsListName: string;
  commentsListName: string;
  commentsReactionsAlias: string;
}

const REACTION_CONFIG: Record<ReactionType, { emoji: string; label: string }> = {
  Like: { emoji: '👍', label: 'Like' },
  Heart: { emoji: '❤️', label: 'Love' },
  Wow: { emoji: '😮', label: 'Wow' },
};

const EMPTY_REACTIONS: IReactionSummary = {
  counts: { Like: 0, Heart: 0, Wow: 0 },
  myReaction: undefined,
  myItemId: undefined,
};

const formatCommentDate = (date: Date): string => {
  const diff = Date.now() - date.getTime();
  const mins = Math.floor(diff / 60000);
  if (mins < 1) return 'Just now';
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(mins / 60);
  if (hrs < 24) return `${hrs}h ago`;
  const days = Math.floor(hrs / 24);
  if (days < 7) return `${days}d ago`;
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
};

const ListAnnouncementDetailsPanel: React.FC<ListAnnouncementDetailsPanelProps> = ({
  announcement, isOpen, onDismiss, onEdit, currentUserId, isEditor, isApprover, hasApproval,
  context, reactionsListName, commentsListName, commentsReactionsAlias,
}) => {
  const [reactions, setReactions] = useState<IReactionSummary>(EMPTY_REACTIONS);
  const [comments, setComments] = useState<IComment[]>([]);
  const [showComments, setShowComments] = useState(false);
  const [newComment, setNewComment] = useState('');
  const [replyText, setReplyText] = useState('');
  const [replyingTo, setReplyingTo] = useState<number | undefined>(undefined);
  const [visibleCount, setVisibleCount] = useState(5);
  const [submittingComment, setSubmittingComment] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const longPressTimer = useRef<number | null>(null);
  const closePickerTimer = useRef<number | null>(null);

  useEffect(() => {
    if (!isOpen || !announcement) {
      setReactions(EMPTY_REACTIONS);
      setComments([]);
      setShowComments(false);
      setNewComment('');
      setReplyText('');
      setReplyingTo(undefined);
      setVisibleCount(5);
      setPickerOpen(false);
      setLoadError(undefined);
      return;
    }
    setLoadError(undefined);
    const reactionSvc = new ReactionService(context, reactionsListName, commentsReactionsAlias);
    const commentSvc = new CommentService(context, commentsListName, commentsReactionsAlias);
    reactionSvc.getReactions(announcement.Id)
      .then(r => setReactions(r))
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        setLoadError(prev => prev ? `${prev} | Reactions: ${msg}` : `Reactions: ${msg}`);
      });
    commentSvc.getComments(announcement.Id)
      .then(c => setComments(c))
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        setLoadError(prev => prev ? `${prev} | Comments: ${msg}` : `Comments: ${msg}`);
      });
  }, [isOpen, announcement?.Id, context, reactionsListName, commentsListName, commentsReactionsAlias]);

  const handleReact = async (type: ReactionType): Promise<void> => {
    if (!announcement) return;
    const svc = new ReactionService(context, reactionsListName, commentsReactionsAlias);
    const prev = reactions;

    if (prev.myReaction === type) {
      const newCounts = { ...prev.counts, [type]: Math.max(0, prev.counts[type] - 1) };
      setReactions({ counts: newCounts, myReaction: undefined, myItemId: undefined });
      setPickerOpen(false);
      await svc.removeReaction(prev);
    } else {
      const newCounts = { ...prev.counts, [type]: prev.counts[type] + 1 };
      if (prev.myReaction) newCounts[prev.myReaction] = Math.max(0, newCounts[prev.myReaction] - 1);
      setReactions({ counts: newCounts, myReaction: type, myItemId: prev.myItemId });
      setPickerOpen(false);
      const updated = await svc.setReaction(announcement.Id, type, prev);
      setReactions(r => ({ ...r, myItemId: updated.myItemId }));
    }
  };

  const handleSubmitComment = async (parentId?: number): Promise<void> => {
    const text = parentId !== undefined ? replyText : newComment;
    if (!text.trim() || submittingComment || !announcement) return;
    setSubmittingComment(true);
    try {
      const svc = new CommentService(context, commentsListName, commentsReactionsAlias);
      const added = await svc.addComment(announcement.Id, text.trim(), context.pageContext.user.displayName, parentId);
      setComments(prev => [...prev, added]);
      if (parentId !== undefined) {
        setReplyText('');
        setReplyingTo(undefined);
      } else {
        setNewComment('');
        setVisibleCount(v => v + 1);
      }
    } catch (err) {
      console.error('Failed to add comment:', err);
    } finally {
      setSubmittingComment(false);
    }
  };

  const handleDeleteComment = async (commentId: number): Promise<void> => {
    try {
      const replyIds = comments
        .filter(c => c.ParentCommentId === commentId)
        .map(c => c.Id);
      const svc = new CommentService(context, commentsListName, commentsReactionsAlias);
      await svc.deleteComment(commentId, replyIds);
      const toRemove = new Set([commentId, ...replyIds]);
      setComments(prev => prev.filter(c => !toRemove.has(c.Id)));
    } catch (err) {
      console.error('Failed to delete comment:', err);
    }
  };

  const onWrapperMouseEnter = (): void => {
    if (closePickerTimer.current) { clearTimeout(closePickerTimer.current); closePickerTimer.current = null; }
    setPickerOpen(true);
  };

  const onWrapperMouseLeave = (): void => {
    closePickerTimer.current = window.setTimeout(() => setPickerOpen(false), 200);
  };

  const onReactPointerDown = (): void => {
    longPressTimer.current = window.setTimeout(() => setPickerOpen(true), 500);
  };

  const onReactPointerUp = (): void => {
    if (longPressTimer.current) { clearTimeout(longPressTimer.current); longPressTimer.current = null; }
  };

  const totalReactions = reactions.counts.Like + reactions.counts.Heart + reactions.counts.Wow;

  if (!isOpen || !announcement) return null; // eslint-disable-line @rushstack/no-new-null

  const isAuthor = announcement.Author?.Id === currentUserId;
  // Without an approval workflow, any Editor/Approver can manage any item — no per-author lock.
  const canEdit = hasApproval ? (isApprover || (isEditor && isAuthor)) : (isEditor || isApprover);

  const getFileIcon = (fileName: string): string => {
    const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
    if (ext === 'pdf') return 'PDF';
    if (ext === 'doc' || ext === 'docx') return 'WordDocument';
    if (ext === 'xls' || ext === 'xlsx' || ext === 'csv') return 'ExcelDocument';
    if (ext === 'ppt' || ext === 'pptx') return 'PowerPointDocument';
    if (ext === 'jpg' || ext === 'jpeg' || ext === 'png' || ext === 'gif' || ext === 'webp') return 'FileImage';
    if (ext === 'mp4') return 'Video';
    if (ext === 'mp3') return 'Music';
    return 'Attach';
  };

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const hasBanner = !!announcement.BannerImageUrl;

  return ReactDOM.createPortal(
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Hero Banner ── */}
        {hasBanner ? (
          <div className={styles.bannerContainer}>
            <img
              src={announcement.BannerImageUrl}
              alt={announcement.Title ?? 'Announcement Banner'}
              className={styles.bannerImage}
            />

            <div className={styles.bannerOverlayContent}>
              <div className={styles.bannerTypeBadgeRow}>
                <span className={styles.bannerTypeBadge}>{announcement.HighlightType ?? 'General'}</span>
                {announcement.Priority === 'Pinned' && (
                  <span className={styles.bannerTypeBadge}>📌 Pinned</span>
                )}
              </div>
              <h1 className={styles.bannerTitle}>{announcement.Title ?? 'Untitled'}</h1>
              {announcement.Author && (
                <span className={styles.bannerAuthor}>
                  <Icon iconName="Contact" /> {announcement.Author.Title}
                </span>
              )}
              {announcement.Created && (
                <span className={styles.bannerPublishDate}>
                  {announcement.Created.toLocaleDateString('en-US', {
                    weekday: 'short', year: 'numeric', month: 'short', day: 'numeric',
                  })}
                </span>
              )}
            </div>

            <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
              <Icon iconName="Cancel" />
            </button>
          </div>
        ) : (
          <button
            className={`${styles.closeBtn} ${styles.closeBtnNoImage}`}
            onClick={onDismiss}
            aria-label="Close"
          >
            <Icon iconName="Cancel" />
          </button>
        )}

        {/* ── Scrollable article body ── */}
        <div className={styles.scrollBody}>
          <div className={styles.content}>

            {!hasBanner && (
              <div className={styles.headerNoBanner}>
                <div className={styles.titleSection}>
                  <div className={styles.typeBadgeRow}>
                    <span className={styles.typeBadge}>{announcement.HighlightType ?? 'General'}</span>
                    {announcement.Priority === 'Pinned' && (
                      <span className={styles.pinnedBadge}>
                        <Icon iconName="Pin" /> Pinned
                      </span>
                    )}
                  </div>
                  <h1 className={styles.announcementTitle}>{announcement.Title ?? 'Untitled'}</h1>
                  {announcement.Author && (
                    <span className={styles.authorLabel}>
                      <Icon iconName="Contact" /> {announcement.Author.Title}
                    </span>
                  )}
                </div>
              </div>
            )}

            {/* Edit / locked controls */}
            <div className={styles.articleActions}>
              {canEdit ? (
                <PrimaryButton
                  text="Edit"
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => onEdit(announcement)}
                  className={styles.editButton}
                />
              ) : (
                <div className={styles.lockedNote}>
                  <Icon iconName="Lock" /> <span>Only the creator can edit</span>
                </div>
              )}
            </div>

            {/* Rejection reason — visible to the author and approvers */}
            {hasApproval && announcement.Status === 'Rejected' && announcement.RejectionReason && (isApprover || isAuthor) && (
              <div className={styles.rejectionNote}>
                <Icon iconName="Warning" /> <span><strong>Rejected:</strong> {announcement.RejectionReason}</span>
              </div>
            )}

            {announcement.Body && (
              <div
                className={styles.bodyBox}
                dangerouslySetInnerHTML={{ __html: announcement.Body }}
              />
            )}

            {announcement.Attachments && announcement.Attachments.length > 0 && (
              <div className={styles.infoCard}>
                <div className={styles.cardHeader}>
                  <Icon iconName="Attach" className={styles.cardIcon} />
                  <h3>Attachments ({announcement.Attachments.length})</h3>
                </div>
                <div className={styles.cardContent}>
                  <div className={styles.attachmentList}>
                    {announcement.Attachments.map((att, idx) => (
                      <a
                        key={idx}
                        href={att.ServerRelativeUrl}
                        target="_blank"
                        rel="noreferrer"
                        className={styles.attachmentCard}
                      >
                        <div className={styles.attachmentIconWrapper}>
                          <Icon iconName={getFileIcon(att.FileName)} className={styles.attachmentTypeIcon} />
                        </div>
                        <div className={styles.attachmentDetails}>
                          <span className={styles.attachmentName}>{att.FileName}</span>
                          <span className={styles.attachmentAction}>Click to open</span>
                        </div>
                        <Icon iconName="OpenInNewWindow" className={styles.attachmentOpenIcon} />
                      </a>
                    ))}
                  </div>
                </div>
              </div>
            )}

            <div className={styles.statusContainer}>
              <span className={`${styles.statusBadge} ${
                announcement.Status === 'Published' ? styles.statusActive :
                announcement.Status === 'Rejected'  ? styles.statusRejected :
                styles.statusDraft
              }`}>
                {announcement.Status ?? 'Draft'}
              </span>
            </div>

            {/* ── Engagement ── */}
            <div className={styles.engagementSection}>

              {loadError && (
                <p style={{ margin: "0 0 8px", fontSize: "0.8em", color: "#d13438" }}>
                  ⚠️ {loadError}
                </p>
              )}

              {(totalReactions > 0 || comments.length > 0) && (
                <div className={styles.engagementMeta}>
                  <div className={styles.reactionTally}>
                    {(["Like", "Heart", "Wow"] as ReactionType[]).map(type =>
                      reactions.counts[type] > 0 ? (
                        <span key={type} className={styles.reactionChip}>
                          {REACTION_CONFIG[type].emoji} {reactions.counts[type]}
                        </span>
                      ) : null
                    )}
                  </div>
                  {comments.length > 0 && (
                    <button className={styles.commentCountBtn} onClick={() => setShowComments(s => !s)}>
                      {comments.length} comment{comments.length !== 1 ? "s" : ""}
                    </button>
                  )}
                </div>
              )}

              <div className={styles.engagementDivider} />

              <div className={styles.engagementActions}>

                <div
                  className={styles.reactWrapper}
                  onMouseEnter={onWrapperMouseEnter}
                  onMouseLeave={onWrapperMouseLeave}
                >
                  {pickerOpen && (
                    <div className={styles.reactionPicker}>
                      {(["Like", "Heart", "Wow"] as ReactionType[]).map(type => (
                        <button
                          key={type}
                          className={`${styles.pickerOption} ${reactions.myReaction === type ? styles.pickerOptionActive : ""}`}
                          onClick={() => handleReact(type).catch(console.error)}
                          title={REACTION_CONFIG[type].label}
                        >
                          <span className={styles.pickerEmoji}>{REACTION_CONFIG[type].emoji}</span>
                          <span className={styles.pickerLabel}>{REACTION_CONFIG[type].label}</span>
                        </button>
                      ))}
                    </div>
                  )}
                  <button
                    className={`${styles.engagementBtn} ${reactions.myReaction ? styles.engagementBtnActive : ""}`}
                    onPointerDown={onReactPointerDown}
                    onPointerUp={onReactPointerUp}
                    onPointerCancel={onReactPointerUp}
                    onClick={reactions.myReaction
                      ? () => handleReact(reactions.myReaction!).catch(console.error)
                      : undefined}
                  >
                    {reactions.myReaction
                      ? `${REACTION_CONFIG[reactions.myReaction].emoji} ${REACTION_CONFIG[reactions.myReaction].label}`
                      : "👍 React"}
                  </button>
                </div>

                <button
                  className={`${styles.engagementBtn} ${showComments ? styles.engagementBtnActive : ""}`}
                  onClick={() => setShowComments(s => !s)}
                >
                  💬 Comment
                </button>

              </div>

              {showComments && (() => {
                const topLevel = comments.filter(c => !c.ParentCommentId);
                const repliesMap = new Map<number, IComment[]>();
                comments
                  .filter(c => c.ParentCommentId !== undefined)
                  .forEach(c => {
                    const existing = repliesMap.get(c.ParentCommentId!) ?? [];
                    repliesMap.set(c.ParentCommentId!, [...existing, c]);
                  });
                const visible = topLevel.slice(0, visibleCount);
                const remaining = topLevel.length - visibleCount;

                return (
                  <div className={styles.commentsSection}>
                    {topLevel.length === 0 && (
                      <p className={styles.noComments}>No comments yet. Be the first!</p>
                    )}

                    {remaining > 0 && (
                      <button
                        className={styles.moreCommentsBtn}
                        onClick={() => setVisibleCount(v => v + 5)}
                      >
                        View {remaining} more comment{remaining !== 1 ? "s" : ""}
                      </button>
                    )}

                    {visible.map(comment => (
                      <div key={comment.Id} className={styles.commentThread}>

                        {/* Main comment */}
                        <div className={styles.commentRow}>
                          <div className={styles.commentAvatar}>
                            {comment.AuthorName.charAt(0).toUpperCase()}
                          </div>
                          <div className={styles.commentBody}>
                            <div className={styles.commentMeta}>
                              <span className={styles.commentAuthor}>{comment.AuthorName}</span>
                              <span className={styles.commentDate}>{formatCommentDate(comment.Created)}</span>
                              {comment.isOwn && (
                                <button
                                  className={styles.commentDeleteBtn}
                                  onClick={() => handleDeleteComment(comment.Id).catch(console.error)}
                                >
                                  Delete
                                </button>
                              )}
                            </div>
                            <p className={styles.commentText}>{comment.CommentText}</p>
                            <button
                              className={styles.replyBtn}
                              onClick={() => setReplyingTo(replyingTo === comment.Id ? undefined : comment.Id)}
                            >
                              {replyingTo === comment.Id ? "Cancel" : "Reply"}
                            </button>
                          </div>
                        </div>

                        {/* Replies */}
                        {(repliesMap.get(comment.Id) ?? []).map(reply => (
                          <div key={reply.Id} className={styles.replyRow}>
                            <span className={styles.replyConnector} aria-hidden="true" />
                            <div className={styles.replyAvatar}>
                              {reply.AuthorName.charAt(0).toUpperCase()}
                            </div>
                            <div className={styles.replyBody}>
                              <div className={styles.commentMeta}>
                                <span className={styles.commentAuthor}>{reply.AuthorName}</span>
                                <span className={styles.commentDate}>{formatCommentDate(reply.Created)}</span>
                                {reply.isOwn && (
                                  <button
                                    className={styles.commentDeleteBtn}
                                    onClick={() => handleDeleteComment(reply.Id).catch(console.error)}
                                  >
                                    Delete
                                  </button>
                                )}
                              </div>
                              <p className={styles.commentText}>{reply.CommentText}</p>
                            </div>
                          </div>
                        ))}

                        {/* Inline reply input */}
                        {replyingTo === comment.Id && (
                          <div className={styles.replyInputRow}>
                            <textarea
                              className={styles.commentTextarea}
                              value={replyText}
                              onChange={e => setReplyText(e.target.value)}
                              placeholder={`Reply to ${comment.AuthorName}…`}
                              rows={2}
                              autoFocus
                              onKeyDown={e => {
                                if (e.key === "Enter" && !e.shiftKey) {
                                  e.preventDefault();
                                  handleSubmitComment(comment.Id).catch(console.error);
                                }
                              }}
                            />
                            <button
                              className={styles.commentSubmitBtn}
                              onClick={() => handleSubmitComment(comment.Id).catch(console.error)}
                              disabled={!replyText.trim() || submittingComment}
                            >
                              {submittingComment ? "…" : "Reply"}
                            </button>
                          </div>
                        )}

                      </div>
                    ))}

                    {/* New top-level comment input */}
                    <div className={styles.commentInputRow}>
                      <textarea
                        className={styles.commentTextarea}
                        value={newComment}
                        onChange={e => setNewComment(e.target.value)}
                        placeholder="Write a comment…"
                        rows={2}
                        onKeyDown={e => {
                          if (e.key === "Enter" && !e.shiftKey) {
                            e.preventDefault();
                            handleSubmitComment().catch(console.error);
                          }
                        }}
                      />
                      <button
                        className={styles.commentSubmitBtn}
                        onClick={() => handleSubmitComment().catch(console.error)}
                        disabled={!newComment.trim() || submittingComment}
                      >
                        {submittingComment ? "…" : "Post"}
                      </button>
                    </div>
                  </div>
                );
              })()}

            </div>

          </div>
        </div>
      </div>
    </div>,
    document.body
  );
};

export default ListAnnouncementDetailsPanel;
