import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Icon, SearchBox, DatePicker, IconButton } from '@fluentui/react';
import { IListAnnouncement } from '../models/IListAnnouncement';
import styles from './ViewAllPanel.module.scss';
import { useState } from 'react';

interface ViewAllPanelProps {
  announcements: IListAnnouncement[];
  isOpen: boolean;
  onDismiss: () => void;
  onSelectAnnouncement: (announcement: IListAnnouncement) => void;
  onAddAnnouncement?: () => void;
  categoryList: string[];
}

const TYPE_COLORS = ['#2e7d32','#1565c0','#6a1b9a','#e65100','#00838f','#ad1457','#c62828','#37474f'];

const ViewAllPanel: React.FC<ViewAllPanelProps> = ({
  announcements, isOpen, onDismiss, onSelectAnnouncement, onAddAnnouncement, categoryList,
}): JSX.Element => {
  const getTypeColor = (type: string): string => {
    const idx = categoryList.findIndex(t => t.toLowerCase() === type?.trim().toLowerCase());
    return idx >= 0 ? TYPE_COLORS[idx % TYPE_COLORS.length] : TYPE_COLORS[0];
  };
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [dateFrom, setDateFrom] = useState<Date | undefined>(undefined);
  const [dateTo, setDateTo] = useState<Date | undefined>(undefined);
  const [showDateFilter, setShowDateFilter] = useState<boolean>(false);

  if (!isOpen) return <></>;

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
  const draftAnnouncements     = filteredAnnouncements.filter(a => a.Status?.trim() === 'Draft');

  const formatDate = (date?: Date): string => {
    if (!date) return 'N/A';
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  const getStatusClass = (status: string): string => {
    if (status === 'Published') return styles.statusActive;
    return styles.statusDraft;
  };


  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const renderAnnouncementCard = (announcement: IListAnnouncement): JSX.Element => (
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
      </div>
    </div>
  );

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
            <span>{filteredAnnouncements.length} announcement{filteredAnnouncements.length !== 1 ? 's' : ''} found</span>
          </div>
        </div>

        {/* ── Scrollable Content ── */}
        <div className={styles.panelContent}>
          {filteredAnnouncements.length === 0 && (
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
        </div>

      </div>
    </div>,
    document.body
  );
};

export default ViewAllPanel;
