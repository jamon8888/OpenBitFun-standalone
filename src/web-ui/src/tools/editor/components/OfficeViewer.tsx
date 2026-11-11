import { useEditorDocument } from '../services/EditorDocument';
import { standaloneEditorFileAccess } from '../services/editorFileAccess';
import type * as XLSX from '@e965/xlsx';
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Download, RotateCw, ZoomIn, ZoomOut } from 'lucide-react';
import { Button, IconButton, Select, Toolbar, ToolbarGroup, ToolbarSeparator, Tooltip } from '@openbitfun/ui';
import { useI18n } from '@/infrastructure/i18n';
import { createLogger } from '@/shared/utils/logger';
import './OfficeViewer.scss';

const log = createLogger('OfficeViewer');

const MIN_ZOOM = 50;
const MAX_ZOOM = 200;
const ZOOM_STEP = 25;
const ZOOM_LEVELS = [50, 75, 100, 125, 150, 200] as const;

/** Rendering caps for spreadsheets; larger sheets stay downloadable via the toolbar. */
const MAX_SHEET_ROWS = 1000;
const MAX_SHEET_COLS = 100;

type OfficeKind = 'docx' | 'xlsx' | 'pptx';

const OFFICE_MIME_TYPES: Record<OfficeKind, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};

/** Legacy .doc/.ppt stay on the code editor; the engines here only parse OOXML and .xls. */
function officeKindFromFileName(fileName: string): OfficeKind | null {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.docx')) return 'docx';
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) return 'xlsx';
  if (lower.endsWith('.pptx')) return 'pptx';
  return null;
}

function decodeBase64(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export interface OfficeViewerProps {
  isActiveTab?: boolean;
  filePath: string;
  /** Owning workspace ID for viewers rendered without an EditorDocument. */
  workspaceId?: string;
  fileName?: string;
  className?: string;
}

export const OfficeViewer: React.FC<OfficeViewerProps> = ({ isActiveTab = true,
  filePath,
  workspaceId,
  fileName,
  className = '',
}) => {
  const documentSession = useEditorDocument();
  const standaloneFiles = useMemo(() => standaloneEditorFileAccess(workspaceId), [workspaceId]);
  const documentFiles = documentSession?.files ?? standaloneFiles;
  const { t, formatNumber } = useI18n('tools');
  const containerRef = useRef<HTMLDivElement | null>(null);
  const zoomRef = useRef<HTMLDivElement | null>(null);
  const docxHostRef = useRef<HTMLDivElement | null>(null);
  const loadEpochRef = useRef(0);
  const bytesRef = useRef<Uint8Array | null>(null);
  const workbookRef = useRef<XLSX.WorkBook | null>(null);
  const [retryKey, setRetryKey] = useState(0);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [unsupported, setUnsupported] = useState(false);
  const [zoom, setZoom] = useState(100);
  const [sheets, setSheets] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState('');
  const [sheetRows, setSheetRows] = useState<string[][]>([]);
  const [truncatedRowCount, setTruncatedRowCount] = useState(0);

  const displayName = useMemo(() => {
    if (fileName && fileName.length > 0) return fileName;
    return filePath.split(/[/\\]/).pop() ?? '';
  }, [fileName, filePath]);
  const kind = useMemo(() => officeKindFromFileName(displayName), [displayName]);

  const renderSheetRows = useCallback(async (workbook: XLSX.WorkBook, sheetName: string) => {
    const XLSX = await import('@e965/xlsx');
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      setSheetRows([]);
      setTruncatedRowCount(0);
      return;
    }
    const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
      header: 1,
      defval: '',
      raw: false,
      blankrows: false,
    });
    const normalized = rows.map(row => row.slice(0, MAX_SHEET_COLS).map(cell => (cell == null ? '' : String(cell))));
    setSheetRows(normalized.slice(0, MAX_SHEET_ROWS));
    setTruncatedRowCount(Math.max(0, normalized.length - MAX_SHEET_ROWS));
  }, []);

  const loadDocument = useCallback(async () => {
    const loadEpoch = ++loadEpochRef.current;
    const currentKind = officeKindFromFileName(displayName);
    setLoading(true);
    setError(null);
    setUnsupported(false);
    setSheets([]);
    setActiveSheet('');
    setSheetRows([]);
    setTruncatedRowCount(0);
    setZoom(100);
    workbookRef.current = null;
    docxHostRef.current?.replaceChildren();

    if (!filePath) {
      setError(t('editor.officeViewer.filePathEmpty'));
      setLoading(false);
      return;
    }
    if (!currentKind) {
      setUnsupported(true);
      setLoading(false);
      return;
    }

    try {
      const encoded = await documentFiles.readFileContent(filePath, 'base64');
      if (loadEpochRef.current !== loadEpoch) return;
      const bytes = decodeBase64(encoded);
      bytesRef.current = bytes;

      if (currentKind === 'pptx') {
        // No open-source pptx engine is bundled; degrade loudly with a download action.
        setUnsupported(true);
        setLoading(false);
        return;
      }

      if (currentKind === 'docx') {
        const { renderAsync } = await import('docx-preview');
        if (loadEpochRef.current !== loadEpoch) return;
        const host = docxHostRef.current;
        if (!host) return;
        host.replaceChildren();
        const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
        await renderAsync(buffer, host, undefined, { inWrapper: true, className: 'openbitfun-office-viewer__docx' });
        if (loadEpochRef.current !== loadEpoch) return;
        setLoading(false);
        return;
      }

      const XLSX = await import('@e965/xlsx');
      if (loadEpochRef.current !== loadEpoch) return;
      const workbook = XLSX.read(bytes, { type: 'array' });
      if (loadEpochRef.current !== loadEpoch) return;
      workbookRef.current = workbook;
      const names = workbook.SheetNames;
      setSheets(names);
      const first = names[0] ?? '';
      setActiveSheet(first);
      await renderSheetRows(workbook, first);
      if (loadEpochRef.current !== loadEpoch) return;
      setLoading(false);
    } catch (loadError) {
      if (loadEpochRef.current !== loadEpoch) return;
      log.warn('Failed to load office document', { filePath, error: errorMessage(loadError) });
      setError(t('editor.officeViewer.loadFailedWithMessage', { message: errorMessage(loadError) }));
      setLoading(false);
    }
  }, [documentFiles, displayName, filePath, renderSheetRows, t]);

  useEffect(() => {
    void loadDocument();
    return () => {
      loadEpochRef.current += 1;
    };
  }, [loadDocument, retryKey]);

  useEffect(() => {
    if (isActiveTab && error && documentSession?.isCurrent()) {
      setRetryKey(current => current + 1);
    }
  }, [documentSession, error, isActiveTab]);

  useEffect(() => {
    if (activeSheet && workbookRef.current) {
      void renderSheetRows(workbookRef.current, activeSheet);
    }
  }, [activeSheet, renderSheetRows]);

  const zoomOut = useCallback(() => {
    setZoom(current => Math.max(MIN_ZOOM, current - ZOOM_STEP));
  }, []);

  const zoomIn = useCallback(() => {
    setZoom(current => Math.min(MAX_ZOOM, current + ZOOM_STEP));
  }, []);

  const selectZoom = useCallback((value: string | number) => {
    const nextZoom = Number(value);
    if (!Number.isFinite(nextZoom)) return;
    setZoom(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, nextZoom)));
  }, []);

  const download = useCallback(() => {
    const bytes = bytesRef.current;
    if (!bytes) return;
    const mime = kind ? OFFICE_MIME_TYPES[kind] : 'application/octet-stream';
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const url = URL.createObjectURL(new Blob([buffer], { type: mime }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = displayName || 'document';
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }, [displayName, kind]);

  const retry = useCallback(() => {
    setRetryKey(current => current + 1);
  }, []);

  const zoomOptions = useMemo(() => ZOOM_LEVELS.map(level => ({
    label: formatNumber(level / 100, { style: 'percent', maximumFractionDigits: 0 }),
    value: level,
  })), [formatNumber]);

  const sheetOptions = useMemo(() => sheets.map(name => ({ label: name, value: name })), [sheets]);
  const headerCells = sheetRows[0] ?? [];
  const bodyRows = sheetRows.slice(1);
  // The content host stays mounted while loading so engines (docx-preview) have a live target;
  // the status overlay is absolutely positioned on top of it.
  const hasContent = !error && !unsupported;

  return (
    <div className={`openbitfun-office-viewer ${className}`}
      data-openbitfun-component="office-viewer"
      data-openbitfun-part="root"
    >
      <Toolbar
        className="openbitfun-office-viewer__toolbar"
        data-openbitfun-component="office-viewer"
        data-openbitfun-part="toolbar"
        leading={(
          <div className="openbitfun-office-viewer__info" data-openbitfun-component="office-viewer" data-openbitfun-part="info">
            {kind === 'xlsx' && sheets.length > 0 && (
              <Select
                aria-label={t('editor.officeViewer.sheet')}
                className="openbitfun-office-viewer__sheet-select"
                size="sm"
                value={activeSheet}
                disabled={loading}
                onValueChange={value => setActiveSheet(String(value))}
                options={sheetOptions}
              />
            )}
          </div>
        )}
        trailing={(
          <>
            <ToolbarGroup>
              <Tooltip content={t('editor.officeViewer.zoomOut')} placement="top">
                <IconButton
                  aria-label={t('editor.officeViewer.zoomOut')}
                  icon={<ZoomOut size={14} />}
                  onClick={zoomOut}
                  disabled={loading || zoom <= MIN_ZOOM}
                  size="sm"
                  variant="quiet"
                />
              </Tooltip>
              <Select
                aria-label={t('editor.officeViewer.zoomLevel')}
                className="openbitfun-office-viewer__zoom-display"
                size="sm"
                value={zoom}
                disabled={loading}
                onValueChange={selectZoom}
                options={zoomOptions}
              />
              <Tooltip content={t('editor.officeViewer.zoomIn')} placement="top">
                <IconButton
                  aria-label={t('editor.officeViewer.zoomIn')}
                  icon={<ZoomIn size={14} />}
                  onClick={zoomIn}
                  disabled={loading || zoom >= MAX_ZOOM}
                  size="sm"
                  variant="quiet"
                />
              </Tooltip>
            </ToolbarGroup>
            <ToolbarSeparator />
            <ToolbarGroup>
              <Tooltip content={t('editor.officeViewer.download')} placement="top">
                <IconButton
                  aria-label={t('editor.officeViewer.download')}
                  icon={<Download size={14} />}
                  onClick={download}
                  disabled={loading || bytesRef.current == null}
                  size="sm"
                  variant="quiet"
                />
              </Tooltip>
            </ToolbarGroup>
          </>
        )}
      />
      <div ref={containerRef} className="openbitfun-office-viewer__container" data-openbitfun-component="office-viewer" data-openbitfun-part="container">
        {hasContent && (
          <div
            ref={zoomRef}
            className="openbitfun-office-viewer__zoomable"
            data-openbitfun-component="office-viewer"
            data-openbitfun-part="content"
            style={{ zoom: zoom / 100 }}
          >
            {kind === 'docx' && (
              <div ref={docxHostRef} className="openbitfun-office-viewer__document" />
            )}
            {kind === 'xlsx' && (
              <div className="openbitfun-office-viewer__sheet-wrap">
                {truncatedRowCount > 0 && (
                  <p className="openbitfun-office-viewer__notice">
                    {t('editor.officeViewer.rowsTruncated', { rows: formatNumber(MAX_SHEET_ROWS) })}
                  </p>
                )}
                {sheetRows.length > 0 && (
                  <table className="openbitfun-office-viewer__sheet">
                    <thead>
                      <tr>{headerCells.map((cell, index) => <th key={index}>{cell}</th>)}</tr>
                    </thead>
                    <tbody>
                      {bodyRows.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          {row.map((cell, cellIndex) => <td key={cellIndex}>{cell}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}
          </div>
        )}
        {loading && (
          <div className="openbitfun-office-viewer__status" data-openbitfun-component="office-viewer" data-openbitfun-part="loading">
            <div className="openbitfun-office-viewer__spinner" />
            <p>{t('editor.officeViewer.loading')}</p>
          </div>
        )}
        {!loading && error && (
          <div className="openbitfun-office-viewer__status openbitfun-office-viewer__status--error" data-openbitfun-component="office-viewer" data-openbitfun-part="error">
            <p>{error}</p>
            <p className="openbitfun-office-viewer__error-path">{filePath}</p>
            <Button size="sm" onClick={retry} leadingIcon={<RotateCw size={14} />}>
              {t('editor.common.retry')}
            </Button>
          </div>
        )}
        {!loading && !error && unsupported && (
          <div className="openbitfun-office-viewer__status" data-openbitfun-component="office-viewer" data-openbitfun-part="unsupported">
            <p>{t('editor.officeViewer.unsupportedFormat', { extension: displayName.split('.').pop() ?? '' })}</p>
            <Button size="sm" onClick={download} leadingIcon={<Download size={14} />}>
              {t('editor.officeViewer.download')}
            </Button>
          </div>
        )}
      </div>
    </div>
  );
};

export default OfficeViewer;
