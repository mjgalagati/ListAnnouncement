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
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/files';
import '@pnp/sp/folders';

type HandlePos = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

const HANDLE_CURSORS: Record<HandlePos, string> = {
  nw: 'nwse-resize', n: 'ns-resize',  ne: 'nesw-resize',
  e:  'ew-resize',   se: 'nwse-resize', s: 'ns-resize',
  sw: 'nesw-resize', w: 'ew-resize',
};

const handleStyle = (pos: HandlePos, r: DOMRect): React.CSSProperties => {
  const S = 9; const h = S / 2;
  const cx = r.left + r.width  / 2 - h;
  const cy = r.top  + r.height / 2 - h;
  const base: React.CSSProperties = {
    position: 'fixed', width: S, height: S,
    background: '#fff', border: '2px solid #171C8F',
    borderRadius: 2, zIndex: 10001, cursor: HANDLE_CURSORS[pos],
    boxSizing: 'border-box',
  };
  switch (pos) {
    case 'nw': return { ...base, left: r.left  - h, top: r.top    - h };
    case 'n':  return { ...base, left: cx,           top: r.top    - h };
    case 'ne': return { ...base, left: r.right - h, top: r.top    - h };
    case 'e':  return { ...base, left: r.right - h, top: cy           };
    case 'se': return { ...base, left: r.right - h, top: r.bottom - h };
    case 's':  return { ...base, left: cx,           top: r.bottom - h };
    case 'sw': return { ...base, left: r.left  - h, top: r.bottom - h };
    case 'w':  return { ...base, left: r.left  - h, top: cy           };
  }
};

interface AddEditListAnnouncementProps {
  isOpen: boolean;
  mode: 'add' | 'edit';
  announcement?: IListAnnouncement;
  context: WebPartContext;
  categoryOptions: string[];
  isApprover: boolean;
  hasApproval: boolean;
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

const APPROVER_STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Draft',     text: 'Draft' },
  { key: 'Published', text: 'Published' },
  { key: 'Rejected',  text: 'Rejected' },
];

// Shown only when an editor reopens a Rejected item — resubmitting is a deliberate,
// separate step from saving edits, so an in-progress edit isn't requeued by accident.
const EDITOR_REJECTED_STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Rejected', text: 'Rejected (keep editing)' },
  { key: 'Draft',    text: 'Draft (resubmit for approval)' },
];

// Without an approval workflow, any Editor/Approver can freely set Draft or Published.
const NO_APPROVAL_STATUS_OPTIONS: IDropdownOption[] = [
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

const MAX_ATTACHMENTS = 3;
const MAX_ATTACHMENT_SIZE_MB = 30;
const MAX_ATTACHMENT_SIZE_BYTES = MAX_ATTACHMENT_SIZE_MB * 1024 * 1024;
const MAX_IMAGE_SIZE_MB = 30;
const MAX_IMAGE_SIZE_BYTES = MAX_IMAGE_SIZE_MB * 1024 * 1024;

const AddEditListAnnouncement: React.FC<AddEditListAnnouncementProps> = ({
  isOpen, mode, announcement, context, categoryOptions, isApprover, hasApproval, onDismiss, onSave,
}) => {
  const bodyEditorRef  = useRef<HTMLDivElement>(null);
  const scrollBodyRef  = useRef<HTMLDivElement>(null);

  const [title, setTitle] = useState('');
  const [category, setCategory] = useState('');
  const [priority, setPriority] = useState<'Pinned' | 'Not Pinned'>('Not Pinned');
  const [status, setStatus] = useState<'Draft' | 'Published' | 'Rejected'>('Draft');
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
  const [attachmentError, setAttachmentError] = useState<string | undefined>(undefined);

  const [selectedImg, setSelectedImg] = useState<HTMLImageElement | null>(null);
  const [imgRect, setImgRect]         = useState<DOMRect | null>(null);

  const categoryDropdownOptions: IDropdownOption[] = categoryOptions.map(t => ({ key: t, text: t }));

  useEffect(() => {
    if (isOpen) {
      if (mode === 'edit' && announcement) {
        setTitle(announcement.Title ?? '');
        setCategory(announcement.HighlightType ?? '');
        setPriority(announcement.Priority === 'Pinned' ? 'Pinned' : 'Not Pinned');
        if (!hasApproval) {
          // No approval workflow — any Editor/Approver can freely set Draft or Published.
          setStatus(announcement.Status === 'Published' ? 'Published' : 'Draft');
        } else if (isApprover) {
          // Approvers get full status control, including publishing directly from the editor.
          setStatus(
            announcement.Status === 'Published' || announcement.Status === 'Rejected'
              ? announcement.Status
              : 'Draft'
          );
        } else {
          // Editors: a Rejected item stays Rejected until they explicitly choose to resubmit.
          setStatus(announcement.Status === 'Rejected' ? 'Rejected' : 'Draft');
        }
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
      setSelectedImg(null); setImgRect(null);
    }
  }, [isOpen, mode, announcement, isApprover, hasApproval]);

  // Keep resize overlay in sync when the scrollable panel body scrolls
  useEffect(() => {
    const el = scrollBodyRef.current;
    if (!el || !selectedImg) return;
    const sync = (): void => setImgRect(selectedImg.getBoundingClientRect());
    el.addEventListener('scroll', sync);
    return () => el.removeEventListener('scroll', sync);
  }, [selectedImg]);

  // Dismiss overlay when clicking outside the editor
  useEffect(() => {
    if (!selectedImg) return;
    const dismiss = (e: MouseEvent): void => {
      if (bodyEditorRef.current?.contains(e.target as Node)) return;
      setSelectedImg(null); setImgRect(null);
    };
    document.addEventListener('mousedown', dismiss);
    return () => document.removeEventListener('mousedown', dismiss);
  }, [selectedImg]);

  // Delete selected image with Delete or Backspace
  useEffect(() => {
    if (!selectedImg) return;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        selectedImg.parentNode?.removeChild(selectedImg);
        setSelectedImg(null); setImgRect(null);
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [selectedImg]);

  if (!isOpen) return null; // eslint-disable-line @rushstack/no-new-null

  const deleteSelectedImg = (): void => {
    if (!selectedImg) return;
    selectedImg.parentNode?.removeChild(selectedImg);
    setSelectedImg(null); setImgRect(null);
  };

  const execFormat = (command: string): void => {
    bodyEditorRef.current?.focus();
    document.execCommand(command, false, undefined);
  };

  const uploadAndInsert = async (file: File, savedRange: Range | null): Promise<void> => {
    const insertNode = (src: string): void => {
      const img = document.createElement('img');
      img.src = src;
      img.style.maxWidth = '100%';
      img.style.height = 'auto';
      img.style.display = 'block';
      if (savedRange) {
        savedRange.deleteContents();
        savedRange.insertNode(img);
        savedRange.setStartAfter(img);
        savedRange.collapse(true);
        const s = window.getSelection();
        if (s) { s.removeAllRanges(); s.addRange(savedRange); }
      } else {
        bodyEditorRef.current?.appendChild(img);
      }
    };

    try {
      const sp = spfi().using(SPFx(context));
      const folder = 'SiteAssets/ListAnnouncementBodyImages';
      try { await sp.web.getFolderByServerRelativePath(folder)(); }
      catch { await sp.web.getFolderByServerRelativePath('SiteAssets').folders.addUsingPath('ListAnnouncementBodyImages'); }
      const ext = file.type.split('/')[1] || 'png';
      const result = await sp.web
        .getFolderByServerRelativePath(folder)
        .files.addUsingPath(`${Date.now()}.${ext}`, file, { Overwrite: true });
      insertNode(result.ServerRelativeUrl);
    } catch (err) {
      console.error('Body image upload failed, using base64 fallback:', err);
      const reader = new FileReader();
      reader.onload = ev => insertNode(ev.target?.result as string);
      reader.readAsDataURL(file);
    }
  };

  const handleBodyPaste = async (e: React.ClipboardEvent<HTMLDivElement>): Promise<void> => {
    const imageItem = Array.from(e.clipboardData.items).find(i => i.type.startsWith('image/'));
    if (!imageItem) return;
    const file = imageItem.getAsFile();
    if (!file) return;
    e.preventDefault();
    if (file.size > MAX_IMAGE_SIZE_BYTES) {
      setAttachmentError(`${file.name || 'Image'} exceeds the ${MAX_IMAGE_SIZE_MB}MB size limit.`);
      return;
    }
    const sel = window.getSelection();
    const savedRange = sel && sel.rangeCount > 0 ? sel.getRangeAt(0).cloneRange() : null;
    await uploadAndInsert(file, savedRange);
  };

  const handleEditorDragOver = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  };

  const handleEditorDrop = async (e: React.DragEvent<HTMLDivElement>): Promise<void> => {
    e.preventDefault();
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
    if (!files.length) return;

    const oversized = files.filter(f => f.size > MAX_IMAGE_SIZE_BYTES);
    if (oversized.length > 0) {
      setAttachmentError(
        `${oversized.map(f => f.name).join(', ')} exceed${oversized.length === 1 ? 's' : ''} the ${MAX_IMAGE_SIZE_MB}MB size limit.`
      );
    }
    const validFiles = files.filter(f => f.size <= MAX_IMAGE_SIZE_BYTES);
    if (!validFiles.length) return;

    let range: Range | null = null;
    if (document.caretRangeFromPoint) {
      range = document.caretRangeFromPoint(e.clientX, e.clientY);
    } else {
      const pos = (document as { caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null }).caretPositionFromPoint?.(e.clientX, e.clientY);
      if (pos) { range = document.createRange(); range.setStart(pos.offsetNode, pos.offset); }
    }

    for (const file of validFiles) {
      await uploadAndInsert(file, range ? range.cloneRange() : null);
    }
  };

  const handleEditorClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    const target = e.target as HTMLElement;
    if (target.tagName === 'IMG') {
      const img = target as HTMLImageElement;
      setSelectedImg(img);
      setImgRect(img.getBoundingClientRect());
    } else {
      setSelectedImg(null); setImgRect(null);
    }
  };

  const startResize = (e: React.MouseEvent, pos: HandlePos): void => {
    e.preventDefault(); e.stopPropagation();
    if (!selectedImg) return;

    const startX   = e.clientX;
    const startY   = e.clientY;
    const startW   = selectedImg.offsetWidth;
    const startH   = selectedImg.offsetHeight;
    const isLeft   = pos.includes('w');
    const isRight  = pos.includes('e');
    const isTop    = pos.includes('n');
    const isBottom = pos.includes('s');
    const onlyVert = pos === 'n' || pos === 's';

    const onMove = (me: MouseEvent): void => {
      const dx = me.clientX - startX;
      const dy = me.clientY - startY;

      if (onlyVert) {
        const newH = Math.max(30, isBottom ? startH + dy : startH - dy);
        selectedImg.style.height = `${Math.round(newH)}px`;
        selectedImg.style.width  = 'auto';
      } else {
        let newW = startW;
        if (isRight) newW = startW + dx;
        if (isLeft)  newW = startW - dx;
        if (!isLeft && !isRight) {
          const newH = Math.max(30, isBottom ? startH + dy : startH - dy);
          selectedImg.style.height = `${Math.round(newH)}px`;
        } else {
          selectedImg.style.width  = `${Math.max(40, Math.round(newW))}px`;
          if (!isTop && !isBottom) selectedImg.style.height = 'auto';
        }
      }
      setImgRect(selectedImg.getBoundingClientRect());
    };

    const onUp = (): void => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      if (selectedImg) setImgRect(selectedImg.getBoundingClientRect());
    };

    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
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
        RejectionReason: status === 'Rejected' ? announcement?.RejectionReason : '',
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
    e.target.value = '';
    if (!file) return;
    if (file.size > MAX_IMAGE_SIZE_BYTES) {
      setAttachmentError(`${file.name} exceeds the ${MAX_IMAGE_SIZE_MB}MB size limit.`);
      return;
    }
    setAttachmentError(undefined);
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
    e.target.value = '';
    if (!files.length) return;

    const oversized = files.filter(f => f.size > MAX_ATTACHMENT_SIZE_BYTES);
    if (oversized.length > 0) {
      setAttachmentError(
        `${oversized.map(f => f.name).join(', ')} exceed${oversized.length === 1 ? 's' : ''} the ${MAX_ATTACHMENT_SIZE_MB}MB size limit.`
      );
      return;
    }

    const currentTotal = existingAttachments.length + newAttachments.length;
    const availableSlots = MAX_ATTACHMENTS - currentTotal;
    if (availableSlots <= 0) {
      setAttachmentError(`You can only attach up to ${MAX_ATTACHMENTS} files.`);
      return;
    }

    if (files.length > availableSlots) {
      setAttachmentError(`Only ${availableSlots} more attachment${availableSlots === 1 ? '' : 's'} can be added (max ${MAX_ATTACHMENTS}).`);
    } else {
      setAttachmentError(undefined);
    }

    setNewAttachments(prev => [...prev, ...files.slice(0, availableSlots)]);
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

  const handles = (['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as HandlePos[]);

  const modal = (
    <>
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
        <div ref={scrollBodyRef} className={styles.panelContent}>
          {saveError && (
            <MessageBar messageBarType={MessageBarType.error} className={styles.errorBar}>
              {saveError}
            </MessageBar>
          )}

          {hasApproval && mode === 'edit' && announcement?.Status === 'Rejected' && announcement.RejectionReason && (
            <MessageBar messageBarType={MessageBarType.severeWarning} className={styles.errorBar}>
              <strong>Rejected:</strong> {announcement.RejectionReason}
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
                <span className={styles.toolbarDivider} />
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Align left"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('justifyLeft'); }}
                >
                  <Icon iconName="AlignLeft" />
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Align center"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('justifyCenter'); }}
                >
                  <Icon iconName="AlignCenter" />
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Align right"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('justifyRight'); }}
                >
                  <Icon iconName="AlignRight" />
                </button>
                <button
                  type="button"
                  className={styles.toolbarBtn}
                  title="Justify"
                  onMouseDown={(e) => { e.preventDefault(); execFormat('justifyFull'); }}
                >
                  <Icon iconName="AlignJustify" />
                </button>
              </div>
              <div
                ref={bodyEditorRef}
                className={styles.bodyEditor}
                contentEditable={true}
                data-placeholder="Write your announcement here…"
                suppressContentEditableWarning={true}
                onPaste={handleBodyPaste}
                onDrop={handleEditorDrop}
                onDragOver={handleEditorDragOver}
                onClick={handleEditorClick}
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
                  <span className={styles.uploadHint}>JPG, PNG, GIF, WEBP supported, up to {MAX_IMAGE_SIZE_MB}MB</span>
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

          {/* Scheduling — with approval on: Approvers get full status control, editors only
              see a status control when resubmitting a Rejected item (else implicitly Draft).
              With approval off: any Editor/Approver can freely set Draft/Published. */}
          {(!hasApproval || isApprover || announcement?.Status === 'Rejected') && (
            <div className={styles.formSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Calendar" className={styles.sectionIcon} />
                <h3>Scheduling</h3>
              </div>
              {!hasApproval ? (
                <Dropdown
                  label="Status"
                  selectedKey={status}
                  options={NO_APPROVAL_STATUS_OPTIONS}
                  onChange={(_, o) => setStatus(o?.key as 'Draft' | 'Published' ?? 'Draft')}
                />
              ) : isApprover ? (
                <Dropdown
                  label="Status"
                  selectedKey={status}
                  options={APPROVER_STATUS_OPTIONS}
                  onChange={(_, o) => setStatus(o?.key as 'Draft' | 'Published' | 'Rejected' ?? 'Draft')}
                />
              ) : (
                <>
                  <Dropdown
                    label="Status"
                    selectedKey={status}
                    options={EDITOR_REJECTED_STATUS_OPTIONS}
                    onChange={(_, o) => setStatus(o?.key as 'Rejected' | 'Draft' ?? 'Rejected')}
                  />
                  <span className={styles.fieldHint}>
                    Editing does not automatically resubmit this for approval — choose &quot;Draft&quot; only when it&apos;s ready to be reviewed again.
                  </span>
                </>
              )}
            </div>
          )}

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
                  <span className={styles.attachmentCount}>{totalAttachments} / {MAX_ATTACHMENTS}</span>
                )}
              </h3>
            </div>

            {attachmentError && (
              <MessageBar
                messageBarType={MessageBarType.error}
                className={styles.errorBar}
                onDismiss={() => setAttachmentError(undefined)}
              >
                {attachmentError}
              </MessageBar>
            )}

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

            {totalAttachments < MAX_ATTACHMENTS ? (
              <label className={styles.attachmentUploadArea} htmlFor="attachment-upload">
                <Icon iconName="Attach" />
                <span>Click to add attachments</span>
                <span className={styles.attachmentHint}>
                  Up to {MAX_ATTACHMENTS} files, {MAX_ATTACHMENT_SIZE_MB}MB each
                </span>
                <input
                  id="attachment-upload"
                  type="file"
                  multiple
                  className={styles.fileInput}
                  onChange={handleAttachmentAdd}
                />
              </label>
            ) : (
              <span className={styles.attachmentHint}>
                Maximum of {MAX_ATTACHMENTS} attachments reached.
              </span>
            )}
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

    {/* ── Image resize overlay ── */}
    {selectedImg && imgRect && (
      <>
        <div style={{
          position: 'fixed',
          left: imgRect.left  - 1, top:    imgRect.top    - 1,
          width: imgRect.width + 2, height: imgRect.height + 2,
          border: '2px solid #171C8F',
          pointerEvents: 'none',
          zIndex: 10000,
          boxSizing: 'border-box',
        }} />
        <button
          onMouseDown={e => { e.preventDefault(); e.stopPropagation(); deleteSelectedImg(); }}
          style={{
            position: 'fixed',
            left: imgRect.right - 12,
            top: imgRect.top - 24,
            width: 22, height: 22,
            background: '#d13438', color: '#fff',
            border: 'none', borderRadius: '50%',
            cursor: 'pointer', fontSize: 13, fontWeight: 700,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            zIndex: 10002, lineHeight: 1,
          }}
          title="Delete image"
        >×</button>
        {handles.map(pos => (
          <div key={pos} style={handleStyle(pos, imgRect)} onMouseDown={e => startResize(e, pos)} />
        ))}
      </>
    )}
    </>
  );

  return ReactDOM.createPortal(modal, document.body);
};

export default AddEditListAnnouncement;
