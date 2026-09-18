import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { PrimaryButton, DefaultButton, Icon, SearchBox, DatePicker, IconButton, TextField } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListAnnouncement } from '../models/IListAnnouncement';
import { ListAnnouncementService } from '../services/ListAnnouncementService';
import styles from './ViewAllPanel.module.scss';
import { useState } from 'react';

const TYPE_COLORS = ['#2e7d32','#1565c0','#6a1b9a','#e65100','#00838f','#ad1457','#c62828','#37474f'];

interface ViewAllPanelProps {
  announcements: IListAnnouncement[];
  isOpen: boolean;
  onDismiss: () => void;
  onSelectAnnouncement: (announcement: IListAnnouncement) => void;
  onAddAnnouncement?: () => void;
  categoryList: string[];
  isEditor: boolean;
  isApprover: boolean;
  hasApproval: boolean;
  currentUserId: number;
  context: WebPartContext;
  sourceList: string;
  onAfterModeration: () => Promise<void>;
}

const ViewAllPanel: React.FC<ViewAllPanelProps> = ({
  announcements, isOpen, onDismiss, onSelectAnnouncement, onAddAnnouncement, categoryList,
  isEditor, isApprover, hasApproval, context, sourceList, onAfterModeration,
}): JSX.Element => {
  const getTypeColor = (type: string): string => {
    const idx = categoryList.findIndex(t => t.toLowerCase() === type?.trim().toLowerCase());
    return idx >= 0 ? TYPE_COLORS[idx % TYPE_COLORS.length] : TYPE_COLORS[0];
  };
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [dateFrom, setDateFrom] = useState<Date | undefined>(undefined);
  const [dateTo, setDateTo] = useState<Date | undefined>(undefined);
  const [showDateFilter, setShowDateFilter] = useState<boolean>(false);
  const [rejectingId, setRejectingId] = useState<number | undefined>(undefined);
  const [rejectReason, setRejectReason] = useState<string>('');
  const [moderatingId, setModeratingId] = useState<number | undefined>(undefined);

  if (!isOpen) return <></>;

  const isModerator = isEditor || isApprover;

  const filteredAnnouncements = announcements.filter(announcement => {
    const matchesSearch =
      announcement.Title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      announcement.HighlightType?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      announcement.Body?.toLowerCase().includes(searchQuery.toLowerCase());

    let matchesDateRange = true;
    if (dateFrom || dateTo) {
      const createdDate = announcement.Created ? new Date(announcement.Created) : new Date();
      if (dateFrom && dateTo) matchesDateRange = createdDate >= dateFrom && createdDate <= dateTo;
      else if (dateFrom) matchesDateRange = createdDate >= dateFrom;
      else if (dateTo) matchesDateRange = createdDate <= dateTo;
    }

    return matchesSearch && matchesDateRange;
  });

  const publishedAnnouncements = filteredAnnouncements.filter(a => a.Status?.trim() === 'Published');
  const draftAnnouncements     = isModerator ? filteredAnnouncements.filter(a => a.Status?.trim() === 'Draft') : [];
  const rejectedAnnouncements  = (isModerator && hasApproval) ? filteredAnnouncements.filter(a => a.Status?.trim() === 'Rejected') : [];

  const visibleCount = publishedAnnouncements.length + draftAnnouncements.length + rejectedAnnouncements.length;

  const formatDate = (date?: Date): string => {
    if (!date) return 'N/A';
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  const getStatusClass = (status: string): string => {
    if (status === 'Published') return styles.statusActive;
    if (status === 'Rejected') return styles.statusRejected;
    return styles.statusDraft;
  };

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const handleApprove = async (announcement: IListAnnouncement): Promise<void> => {
    setModeratingId(announcement.Id);
    try {
      const service = new ListAnnouncementService(context, sourceList);
      await service.approveAnnouncement(announcement.Id);
      await onAfterModeration();
    } catch (err) {
      console.error('Failed to publish announcement:', err);
    } finally {
      setModeratingId(undefined);
    }
  };

  const startReject = (announcementId: number): void => {
    setRejectingId(announcementId);
    setRejectReason('');
  };

  const cancelReject = (): void => {
    setRejectingId(undefined);
    setRejectReason('');
  };

  const confirmReject = async (announcement: IListAnnouncement): Promise<void> => {
    if (!rejectReason.trim()) return;
    setModeratingId(announcement.Id);
    try {
      const service = new ListAnnouncementService(context, sourceList);
      await service.rejectAnnouncement(announcement.Id, rejectReason.trim());
      await onAfterModeration();
      cancelReject();
    } catch (err) {
      console.error('Failed to reject announcement:', err);
    } finally {
      setModeratingId(undefined);
    }
  };

  const renderAnnouncementCard = (announcement: IListAnnouncement): JSX.Element => {
    const canModerate = hasApproval && isApprover && announcement.Status?.trim() === 'Draft';
    const isRejecting = rejectingId === announcement.Id;
    const isBusy = moderatingId === announcement.Id;

    return (
      <div
        key={announcement.Id}
        className={styles.announcementCard}
        onClick={() => onSelectAnnouncement(announcement)}
      >
        <div className={styles.cardContent}>
          <div className={styles.cardHeader}>
            <h3 className={styles.cardTitle}>{announcement.Title}</h3>
            <div className={styles.cardBadges}>
              {announcement.Priority === 'Pinned' && (
                <span className={styles.pinnedBadge}>📌 Pinned</span>
              )}
              <span className={`${styles.statusBadge} ${getStatusClass(announcement.Status)}`}>
                {announcement.Status}
              </span>
            </div>
          </div>
          <div className={styles.cardInfo}>
            <div className={styles.infoItem}>
              <Icon iconName="Calendar" className={styles.infoIcon} />
              <span>{formatDate(announcement.Created)}</span>
            </div>
            <div className={styles.infoItem}>
              <Icon iconName="Tag" className={styles.infoIcon} />
              <span className={styles.typeTag} style={{ color: getTypeColor(announcement.HighlightType) }}>{announcement.HighlightType}</span>
            </div>
            {announcement.TargetAudienceType && announcement.TargetAudienceType !== 'All' && (
              <div className={styles.infoItem}>
                <Icon iconName="People" className={styles.infoIcon} />
                <span>{announcement.TargetAudienceType === 'Specific' ? 'Specific Audience' : 'All Except'}</span>
              </div>
            )}
          </div>

          {canModerate && !isRejecting && (
            <div className={styles.moderationActions} onClick={e => e.stopPropagation()}>
              <PrimaryButton
                text="Publish"
                iconProps={{ iconName: 'CheckMark' }}
                onClick={() => handleApprove(announcement)}
                disabled={isBusy}
                className={styles.approveBtn}
              />
              <DefaultButton
                text="Reject"
                iconProps={{ iconName: 'Cancel' }}
                onClick={() => startReject(announcement.Id)}
                disabled={isBusy}
                className={styles.rejectBtn}
              />
            </div>
          )}

          {canModerate && isRejecting && (
            <div className={styles.rejectForm} onClick={e => e.stopPropagation()}>
              <TextField
                placeholder="Reason for rejection..."
                value={rejectReason}
                onChange={(_, v) => setRejectReason(v || '')}
                multiline
                rows={2}
                autoFocus
              />
              <div className={styles.rejectFormActions}>
                <PrimaryButton
                  text="Confirm Reject"
                  onClick={() => confirmReject(announcement)}
                  disabled={!rejectReason.trim() || isBusy}
                  className={styles.rejectBtn}
                />
                <DefaultButton text="Cancel" onClick={cancelReject} disabled={isBusy} />
              </div>
            </div>
          )}
        </div>
      </div>
    );
  };

  return ReactDOM.createPortal(
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Header ── */}
        <div className={styles.panelHeader}>
          <div className={styles.headerTop}>
            <h2 className={styles.panelTitle}>All Announcements</h2>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {onAddAnnouncement && (
                <button className={styles.addButton} onClick={onAddAnnouncement}>
                  <Icon iconName="Add" /> Add Announcement
                </button>
              )}
              <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
                <Icon iconName="Cancel" />
              </button>
            </div>
          </div>

          <div className={styles.filterBar}>
            <SearchBox
              placeholder="Search announcements..."
              onChange={(_, newValue) => setSearchQuery(newValue || '')}
              className={styles.searchBox}
            />
            <IconButton
              iconProps={{ iconName: showDateFilter ? 'ChevronUp' : 'Calendar' }}
              onClick={() => setShowDateFilter(!showDateFilter)}
              className={styles.dateFilterToggle}
              toggle
              checked={showDateFilter}
            />
          </div>

          {showDateFilter && (
            <div className={styles.dateFilters}>
              <div className={styles.datePickerWrapper}>
                <label className={styles.dateLabel}>From:</label>
                <DatePicker
                  placeholder="Select start date"
                  value={dateFrom}
                  onSelectDate={(date) => setDateFrom(date || undefined)}
                  formatDate={(date) => date?.toLocaleDateString() || ''}
                  className={styles.datePicker}
                />
                {dateFrom && (
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setDateFrom(undefined)}
                    className={styles.clearButton}
                  />
                )}
              </div>
              <div className={styles.datePickerWrapper}>
                <label className={styles.dateLabel}>To:</label>
                <DatePicker
                  placeholder="Select end date"
                  value={dateTo}
                  onSelectDate={(date) => setDateTo(date || undefined)}
                  formatDate={(date) => date?.toLocaleDateString() || ''}
                  minDate={dateFrom}
                  className={styles.datePicker}
                />
                {dateTo && (
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setDateTo(undefined)}
                    className={styles.clearButton}
                  />
                )}
              </div>
              {(dateFrom || dateTo) && (
                <IconButton
                  iconProps={{ iconName: 'ClearFilter' }}
                  onClick={() => { setDateFrom(undefined); setDateTo(undefined); }}
                  className={styles.clearAllButton}
                />
              )}
            </div>
          )}

          <div className={styles.announcementCount}>
            <span>{visibleCount} announcement{visibleCount !== 1 ? 's' : ''} found</span>
          </div>
        </div>

        {/* ── Scrollable Content ── */}
        <div className={styles.panelContent}>
          {visibleCount === 0 && (
            <div className={styles.emptyState}>
              <Icon iconName="Megaphone" className={styles.emptyIcon} />
              <p>No announcements found</p>
              <span>Try adjusting your search or add a new announcement</span>
            </div>
          )}

          {publishedAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="MegaphoneSolid" className={styles.sectionIcon} />
                <h3>Published ({publishedAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {publishedAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {draftAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Edit" className={styles.sectionIcon} />
                <h3>Draft ({draftAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {draftAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}

          {rejectedAnnouncements.length > 0 && (
            <div className={styles.announcementSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="ErrorBadge" className={styles.sectionIcon} />
                <h3>Rejected ({rejectedAnnouncements.length})</h3>
              </div>
              <div className={styles.announcementGrid}>
                {rejectedAnnouncements.map(renderAnnouncementCard)}
              </div>
            </div>
          )}
        </div>

      </div>
    </div>,
    document.body
  );
};

export default ViewAllPanel;
