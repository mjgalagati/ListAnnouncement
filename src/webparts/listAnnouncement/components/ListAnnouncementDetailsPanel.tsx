import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, Icon } from '@fluentui/react';
import { IListAnnouncement } from '../models/IListAnnouncement';
import styles from './ListAnnouncementDetailsPanel.module.scss';

interface ListAnnouncementDetailsPanelProps {
  announcement?: IListAnnouncement;
  isOpen: boolean;
  onDismiss: () => void;
  onEdit: (announcement: IListAnnouncement) => void;
  currentUserId: number;
  isAdmin: boolean;
}

const ListAnnouncementDetailsPanel: React.FC<ListAnnouncementDetailsPanelProps> = ({
  announcement, isOpen, onDismiss, onEdit, currentUserId, isAdmin,
}) => {
  if (!isOpen || !announcement) return null; // eslint-disable-line @rushstack/no-new-null

  const isAuthor = announcement.Author?.Id === currentUserId;
  const canEdit = isAdmin || isAuthor;

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
                  <Icon iconName="Lock" /> <span>Only editors and the creator can edit</span>
                </div>
              )}
            </div>

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
                announcement.Status === 'Published' ? styles.statusActive : styles.statusDraft
              }`}>
                {announcement.Status ?? 'Draft'}
              </span>
            </div>

          </div>
        </div>
      </div>
    </div>,
    document.body
  );
};

export default ListAnnouncementDetailsPanel;
