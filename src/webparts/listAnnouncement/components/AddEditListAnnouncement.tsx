import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import * as ReactDOM from 'react-dom';
import {
  PrimaryButton, DefaultButton,
  TextField, Dropdown, IDropdownOption, Icon,
  Spinner, SpinnerSize, MessageBar, MessageBarType,
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IListAnnouncement, IListAnnouncementAttachment, IListAnnouncementAudience } from '../models/IListAnnouncement';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './AddEditListAnnouncement.module.scss';

interface AddEditListAnnouncementProps {
  isOpen: boolean;
  mode: 'add' | 'edit';
  announcement?: IListAnnouncement;
  context: WebPartContext;
  categoryOptions: string[];
  onDismiss: () => void;
  onSave: (
    data: Partial<IListAnnouncement>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ) => Promise<void>;
}

const PRIORITY_OPTIONS: IDropdownOption[] = [
  { key: 'Not Pinned', text: 'Not Pinned' },
  { key: 'Pinned',     text: 'Pinned' },
];

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Draft',     text: 'Draft' },
  { key: 'Published', text: 'Published' },
];

const AUDIENCE_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'All',      text: 'All' },
  { key: 'Specific', text: 'Specific' },
  { key: 'Except',   text: 'Except' },
];

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

const formatFileSize = (bytes: number): string => {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
};

const AddEditListAnnouncement: React.FC<AddEditListAnnouncementProps> = ({
  isOpen, mode, announcement, context, categoryOptions, onDismiss, onSave,
}) => {
  const bodyEditorRef = useRef<HTMLDivElement>(null);

  const [title, setTitle] = useState('');
  const [category, setCategory] = useState('');
  const [priority, setPriority] = useState<'Pinned' | 'Not Pinned'>('Not Pinned');
  const [status, setStatus] = useState<'Draft' | 'Published'>('Draft');
  const [targetAudienceType, setTargetAudienceType] = useState<'All' | 'Specific' | 'Except'>('All');
  const [targetAudience, setTargetAudience] = useState<IListAnnouncementAudience[]>([]);
  const [bannerFile, setBannerFile] = useState<File | undefined>(undefined);
  const [bannerPreview, setBannerPreview] = useState<string | undefined>(undefined);
  const [newAttachments, setNewAttachments] = useState<File[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<IListAnnouncementAttachment[]>([]);
  const [deletedAttachmentNames, setDeletedAttachmentNames] = useState<string[]>([]);
  const [isSaving, setIsSaving] = useState(false);
  const [errors, setErrors] = useState<{ [key: string]: string }>({});
  const [saveError, setSaveError] = useState<string | undefined>(undefined);

  const categoryDropdownOptions: IDropdownOption[] = categoryOptions.map(t => ({ key: t, text: t }));

  useEffect(() => {
    if (isOpen) {
      if (mode === 'edit' && announcement) {
        setTitle(announcement.Title ?? '');
        setCategory(announcement.HighlightType ?? '');
        setPriority(announcement.Priority === 'Pinned' ? 'Pinned' : 'Not Pinned');
        setStatus((announcement.Status === 'Draft' || announcement.Status === 'Published') ? announcement.Status : 'Draft');
        setTargetAudienceType(announcement.TargetAudienceType ?? 'All');
        setTargetAudience(announcement.TargetAudience ?? []);
        setBannerPreview(announcement.BannerImageUrl ?? undefined);
        setExistingAttachments(announcement.Attachments ?? []);
        if (bodyEditorRef.current) bodyEditorRef.current.innerHTML = announcement.Body ?? '';
      } else {
        setTitle('');
        setCategory(categoryOptions[0] ?? '');
        setPriority('Not Pinned');
        setStatus('Draft');
        setTargetAudienceType('All');
        setTargetAudience([]);
        setBannerPreview(undefined);
        setExistingAttachments([]);
        if (bodyEditorRef.current) bodyEditorRef.current.innerHTML = '';
      }
      setBannerFile(undefined);
      setNewAttachments([]);
      setDeletedAttachmentNames([]);
      setErrors({});
      setSaveError(undefined);
    }
  }, [isOpen, mode, announcement]);

  if (!isOpen) return null; // eslint-disable-line @rushstack/no-new-null

  const execFormat = (command: string): void => {
    document.execCommand(command, false, undefined);
  };

  const validate = (): boolean => {
    const newErrors: { [key: string]: string } = {};
    if (!title.trim()) newErrors.title = 'Title is required';
    if (!category) newErrors.category = 'Category is required';
    if (targetAudienceType !== 'All' && targetAudience.length === 0) {
      newErrors.targetAudience = 'Please select at least one person for the target audience';
    }
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSave = async (): Promise<void> => {
    if (!validate()) return;
    setIsSaving(true);
    setSaveError(undefined);
    try {
      const data: Partial<IListAnnouncement> = {
        Title: title,
        Body: bodyEditorRef.current?.innerHTML ?? '',
        HighlightType: category,
        Site: announcement?.Site ?? 'All',
        Priority: priority,
        Status: status,
        TargetAudienceType: targetAudienceType,
        TargetAudience: targetAudience,
        BannerImageUrl: bannerPreview,
      };
      await onSave(data, bannerFile, newAttachments, deletedAttachmentNames);
    } catch (err) {
      console.error('Save error:', err);
      setSaveError('Failed to save announcement. Please try again.');
    } finally {
      setIsSaving(false);
    }
  };

  const handleBannerChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBannerFile(file);
    const reader = new FileReader();
    reader.onload = (ev) => setBannerPreview(ev.target?.result as string);
    reader.readAsDataURL(file);
  };

  const handleRemoveBanner = (): void => {
    setBannerFile(undefined);
    setBannerPreview(undefined);
  };

  const handleAttachmentAdd = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const files = Array.from(e.target.files ?? []);
    if (!files.length) return;
    setNewAttachments(prev => [...prev, ...files]);
  };

  const handleRemoveNewAttachment = (index: number): void => {
    setNewAttachments(prev => prev.filter((_, i) => i !== index));
  };

  const handleRemoveExistingAttachment = (fileName: string): void => {
    setExistingAttachments(prev => prev.filter(a => a.FileName !== fileName));
    setDeletedAttachmentNames(prev => [...prev, fileName]);
  };

  const totalAttachments = existingAttachments.length + newAttachments.length;

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const modal = (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Header ── */}
        <div className={styles.panelHeader}>
          <h2 className={styles.panelTitle}>
            <Icon iconName={mode === 'add' ? 'Add' : 'Edit'} className={styles.titleIcon} />
            {mode === 'add' ? 'New Announcement' : 'Edit Announcement'}
          </h2>
          <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
            <Icon iconName="Cancel" />
          </button>
        </div>

        {/* ── Scrollable Form Body ── */}
        <div className={styles.panelContent}>
          {saveError && (
            <MessageBar messageBarType={MessageBarType.error} className={styles.errorBar}>
              {saveError}
            </MessageBar>
          )}

          {/* Basic Info */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Info" className={styles.sectionIcon} />
              <h3>Basic Information</h3>
            </div>
            <TextField
              label="Title"
              required
              value={title}
              onChange={(_, v) => setTitle(v ?? '')}
              errorMessage={errors.title}
              placeholder="Enter announcement title"
            />
            <Dropdown
              label="Category"
              required
              selectedKey={category}
              options={categoryDropdownOptions}
              onChange={(_, o) => setCategory(o?.key as string ?? '')}
              errorMessage={errors.category}
              disabled={categoryDropdownOptions.length === 0}
              placeholder={categoryDropdownOptions.length === 0 ? 'Configure Categories in web part settings' : 'Select category'}
            />
          </div>

          {/* Body */}
          <div className={styles.formSection}>
            <div className={styles.bodyField}>
              <label className={styles.bodyLabel}>Body</label>
              <div className={styles.bodyToolbar}>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Bold"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('bold'); }}
                >
                  <b>B</b>
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Italic"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('italic'); }}
                >
                  <i>I</i>
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Underline"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('underline'); }}
                >
                  <u>U</u>
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Bullet list"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('insertUnorderedList'); }}
                >
                  <Icon iconName="BulletedList" />
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Numbered list"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('insertOrderedList'); }}
                >
                  <Icon iconName="NumberedList" />
                </button>
              </div>
              <div
                ref={bodyEditorRef}
                className={styles.bodyEditor}
                contentEditable={true}
                data-placeholder="Write your announcement here…"
                suppressContentEditableWarning={true}
              />
            </div>
          </div>

          {/* Banner Image */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Photo2" className={styles.sectionIcon} />
              <h3>Banner Image</h3>
            </div>
            <div className={styles.bannerUpload}>
              {bannerPreview ? (
                <div className={styles.bannerPreviewWrapper}>
                  <img src={bannerPreview} alt="Banner preview" className={styles.bannerPreview} />
                  <DefaultButton
                    text="Remove Banner"
                    iconProps={{ iconName: 'Delete' }}
                    onClick={handleRemoveBanner}
                    className={styles.removeBtn}
                  />
                </div>
              ) : (
                <label className={styles.uploadArea} htmlFor="banner-upload">
                  <Icon iconName="Photo2" className={styles.uploadIcon} />
                  <span className={styles.uploadLabel}>Click to upload banner image</span>
                  <span className={styles.uploadHint}>JPG, PNG, GIF, WEBP supported</span>
                  <input
                    id="banner-upload"
                    type="file"
                    accept="image/*"
                    className={styles.fileInput}
                    onChange={handleBannerChange}
                  />
                </label>
              )}
            </div>
          </div>

          {/* Pin */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Pin" className={styles.sectionIcon} />
              <h3>Pin to Top</h3>
            </div>
            <Dropdown
              label="Pinned"
              selectedKey={priority}
              options={PRIORITY_OPTIONS}
              onChange={(_, o) => setPriority(o?.key === 'Pinned' ? 'Pinned' : 'Not Pinned')}
            />
          </div>

          {/* Scheduling */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Calendar" className={styles.sectionIcon} />
              <h3>Scheduling</h3>
            </div>
            <Dropdown
              label="Status"
              selectedKey={status}
              options={STATUS_OPTIONS}
              onChange={(_, o) => setStatus(o?.key as 'Draft' | 'Published' ?? 'Draft')}
            />
          </div>

          {/* Target Audience */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="People" className={styles.sectionIcon} />
              <h3>Target Audience</h3>
            </div>
            <Dropdown
              label="Audience Type"
              selectedKey={targetAudienceType}
              options={AUDIENCE_TYPE_OPTIONS}
              onChange={(_, o) => {
                setTargetAudienceType(o?.key as 'All' | 'Specific' | 'Except' ?? 'All');
                setTargetAudience([]);
              }}
            />
            {targetAudienceType === 'Specific' && (
              <span className={styles.fieldHint}>
                Only the selected people below will see this announcement.
              </span>
            )}
            {targetAudienceType === 'Except' && (
              <span className={styles.fieldHint}>
                Everyone will see this announcement <strong>except</strong> the selected people below.
              </span>
            )}
            {targetAudienceType !== 'All' && (
              <>
                <PeoplePicker
                  context={context as any} // eslint-disable-line @typescript-eslint/no-explicit-any
                  titleText="Select People"
                  personSelectionLimit={50}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  defaultSelectedUsers={targetAudience.map(p => p.Title)}
                  onChange={(items) => {
                    const people: IListAnnouncementAudience[] = (items || []).map((item: any) => ({ // eslint-disable-line @typescript-eslint/no-explicit-any
                      Id: parseInt(item.id, 10),
                      Title: item.text,
                    }));
                    setTargetAudience(people);
                  }}
                />
                {errors.targetAudience && <span className={styles.error}>{errors.targetAudience}</span>}
              </>
            )}
          </div>

          {/* Attachments */}
          <div className={styles.formSection}>
            <div className={styles.sectionHeader}>
              <Icon iconName="Attach" className={styles.sectionIcon} />
              <h3>
                Attachments
                {totalAttachments > 0 && (
                  <span className={styles.attachmentCount}>{totalAttachments}</span>
                )}
              </h3>
            </div>

            {existingAttachments.length > 0 && (
              <div className={styles.attachmentList}>
                {existingAttachments.map((att) => (
                  <div key={att.FileName} className={styles.attachmentItem}>
                    <div className={styles.attachmentInfo}>
                      <Icon iconName={getFileIcon(att.FileName)} className={styles.attachmentIcon} />
                      <span className={styles.attachmentName}>{att.FileName}</span>
                    </div>
                    <div className={styles.attachmentActions}>
                      <a href={att.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.attachmentDownload}>
                        <Icon iconName="Download" />
                      </a>
                      <button className={styles.attachmentRemove} onClick={() => handleRemoveExistingAttachment(att.FileName)}>
                        <Icon iconName="Delete" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {newAttachments.length > 0 && (
              <div className={styles.attachmentList}>
                {newAttachments.map((file, idx) => (
                  <div key={idx} className={`${styles.attachmentItem} ${styles.attachmentNew}`}>
                    <div className={styles.attachmentInfo}>
                      <Icon iconName={getFileIcon(file.name)} className={styles.attachmentIcon} />
                      <span className={styles.attachmentName}>{file.name}</span>
                      <span className={styles.attachmentSize}>{formatFileSize(file.size)}</span>
                    </div>
                    <div className={styles.attachmentActions}>
                      <button className={styles.attachmentRemove} onClick={() => handleRemoveNewAttachment(idx)}>
                        <Icon iconName="Delete" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            <label className={styles.attachmentUploadArea} htmlFor="attachment-upload">
              <Icon iconName="Attach" />
              <span>Click to add attachments</span>
              <input
                id="attachment-upload"
                type="file"
                multiple
                className={styles.fileInput}
                onChange={handleAttachmentAdd}
              />
            </label>
          </div>
        </div>

        {/* ── Sticky Action Buttons ── */}
        <div className={styles.actionButtons}>
          <DefaultButton text="Cancel" onClick={onDismiss} disabled={isSaving} />
          <PrimaryButton
            text={isSaving ? 'Saving...' : mode === 'add' ? 'Add Announcement' : 'Save Changes'}
            onClick={handleSave}
            disabled={isSaving}
          >
            {isSaving && <Spinner size={SpinnerSize.small} />}
          </PrimaryButton>
        </div>

      </div>
    </div>
  );

  return ReactDOM.createPortal(modal, document.body);
};

export default AddEditListAnnouncement;
