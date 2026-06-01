import * as React from "react";
import { useEffect, useState } from "react";
import { Icon } from "@fluentui/react";
import styles from "./ListAnnouncements.module.scss";
import { IListAnnouncementsProps } from "./IListAnnouncementsProps";
import { ListAnnouncementService } from "../services/ListAnnouncementService";
import { IListAnnouncement } from "../models/IListAnnouncement";
import ListAnnouncementDetailsPanel from "./ListAnnouncementDetailsPanel";
import ViewAllPanel from "./ViewAllPanel";
import AddEditListAnnouncement from "./AddEditListAnnouncement";


const TYPE_COLORS = [
  '#2e7d32',
  '#1565c0',
  '#6a1b9a',
  '#e65100',
  '#00838f',
  '#ad1457',
  '#c62828',
  '#37474f',
];

const stripHtml = (html: string): string =>
  html.replace(/<[^>]+>/g, '').trim();

const ListAnnouncements = (props: IListAnnouncementsProps): JSX.Element => {
  const [announcements, setAnnouncements] = useState<IListAnnouncement[]>([]);
  const [selectedAnnouncement, setSelectedAnnouncement] = useState<IListAnnouncement | undefined>(undefined);
  const [isDetailsOpen, setIsDetailsOpen] = useState(false);
  const [isViewAllOpen, setIsViewAllOpen] = useState(false);
  const [isAddEditOpen, setIsAddEditOpen] = useState(false);
  const [addEditMode, setAddEditMode] = useState<'add' | 'edit'>('add');
  const [announcementToEdit, setAnnouncementToEdit] = useState<IListAnnouncement | undefined>(undefined);

  const categoryList = props.highlightTypes
    ? props.highlightTypes.split(',').map(t => t.trim()).filter(Boolean)
    : [];

  const getTypeColor = (type: string): string => {
    const idx = categoryList.findIndex(t => t.toLowerCase() === type?.trim().toLowerCase());
    return idx >= 0 ? TYPE_COLORS[idx % TYPE_COLORS.length] : TYPE_COLORS[0];
  };

  const loadAnnouncements = async (): Promise<void> => {
    try {
      const service = new ListAnnouncementService(props.context, props.sourceList);
      const data = await service.getListAnnouncements();
      setAnnouncements(data);
    } catch (err) {
      console.error('Failed to load announcements', err);
    }
  };

  useEffect(() => {
    if (!props.sourceList) return;
    loadAnnouncements().catch(console.error);
  }, [props.sourceList, props.context]);

  const visibleAnnouncements = announcements
    .filter(a => a.Status?.trim() === 'Published')
    .slice(0, props.itemLimit ?? 5);

  const formatDate = (date: Date): string =>
    date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const openDetails = (announcement: IListAnnouncement): void => {
    setSelectedAnnouncement(announcement);
    setIsDetailsOpen(true);
  };

  const closeDetails = (): void => {
    setIsDetailsOpen(false);
    setSelectedAnnouncement(undefined);
  };

  const openEdit = (announcement: IListAnnouncement): void => {
    setIsDetailsOpen(false);
    setAddEditMode('edit');
    setAnnouncementToEdit(announcement);
    setIsAddEditOpen(true);
  };

  const openAdd = (): void => {
    setAddEditMode('add');
    setAnnouncementToEdit(undefined);
    setIsAddEditOpen(true);
  };

  const closeAddEdit = (): void => {
    setIsAddEditOpen(false);
    if (announcementToEdit) {
      setSelectedAnnouncement(announcementToEdit);
      setIsDetailsOpen(true);
    }
    setAnnouncementToEdit(undefined);
  };

  const handleSave = async (
    data: Partial<IListAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> => {
    const service = new ListAnnouncementService(props.context, props.sourceList);
    if (addEditMode === 'add') {
      await service.addListAnnouncement(data, bannerFile, attachments);
    } else if (addEditMode === 'edit' && announcementToEdit) {
      await service.updateListAnnouncement(announcementToEdit.Id, data, bannerFile, attachments, deletedAttachmentNames);
    }
    setIsAddEditOpen(false);
    setAnnouncementToEdit(undefined);
    setSelectedAnnouncement(undefined);
    setIsDetailsOpen(false);
    setIsViewAllOpen(false);
    await loadAnnouncements();
  };

  const renderPanels = (): JSX.Element => (
    <>
      <ViewAllPanel
        announcements={announcements}
        isOpen={isViewAllOpen}
        onDismiss={() => setIsViewAllOpen(false)}
        onSelectAnnouncement={openDetails}
        onAddAnnouncement={props.isAdmin ? openAdd : undefined}
        categoryList={categoryList}
      />
      <ListAnnouncementDetailsPanel
        announcement={selectedAnnouncement}
        isOpen={isDetailsOpen}
        onDismiss={closeDetails}
        onEdit={openEdit}
        currentUserId={props.currentUserId}
        isAdmin={props.isAdmin}
      />
      <AddEditListAnnouncement
        isOpen={isAddEditOpen}
        mode={addEditMode}
        announcement={announcementToEdit}
        context={props.context}
        categoryOptions={categoryList}
        onDismiss={closeAddEdit}
        onSave={handleSave}
      />
    </>
  );

  if (!props.sourceList) {
    return (
      <div className={styles.container}>
        <h2 className={styles.webpartTitle}>{props.webpartTitle || 'ESH Bulletin'}</h2>
        <p className={styles.noList}>Please select a list in the web part settings to get started.</p>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.headerRow}>
        <h2 className={styles.webpartTitle}>{props.webpartTitle || 'ESH Bulletin'}</h2>
        <div className={styles.headerActions}>
          {props.isAdmin && (
            <button className={styles.addBtn} onClick={openAdd}>
              <Icon iconName="Add" /> Add
            </button>
          )}
          <button className={styles.viewAllBtn} onClick={() => setIsViewAllOpen(true)}>
            View all
          </button>
        </div>
      </div>

      {visibleAnnouncements.length === 0 ? (
        <div className={styles.emptyState}>
          <p>No announcements published yet.</p>
        </div>
      ) : (
        <ul className={styles.itemsList}>
          {visibleAnnouncements.map(announcement => {
            const bodyText = stripHtml(announcement.Body ?? '');
            return (
              <li key={announcement.Id} className={styles.item} onClick={() => openDetails(announcement)}>
                <div className={styles.itemThumb}>
                  {announcement.BannerImageUrl ? (
                    <img src={announcement.BannerImageUrl} alt="" className={styles.thumbImg} />
                  ) : (
                    <Icon iconName="Megaphone" />
                  )}
                </div>
                <div className={styles.itemBody}>
                  <div
                    className={styles.itemCategory}
                    style={{ color: getTypeColor(announcement.HighlightType) }}
                  >
                    {announcement.HighlightType}
                  </div>
                  <div className={styles.itemTitle}>{announcement.Title}</div>
                  {bodyText && <div className={styles.itemDesc}>{bodyText}</div>}
                </div>
                <div className={styles.itemDate}>{announcement.Created ? formatDate(announcement.Created) : ''}</div>
              </li>
            );
          })}
        </ul>
      )}

      {renderPanels()}
    </div>
  );
};

export default ListAnnouncements;
