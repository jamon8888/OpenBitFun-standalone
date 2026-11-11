// @vitest-environment jsdom

import React, { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { OfficeViewer } from './OfficeViewer';

const mocks = vi.hoisted(() => ({
  readFileContent: vi.fn(),
  renderDocx: vi.fn(),
  xlsxRead: vi.fn(),
  sheetToJson: vi.fn(),
}));

vi.mock('docx-preview', () => ({
  renderAsync: (...args: unknown[]) => mocks.renderDocx(...args),
}));

vi.mock('@e965/xlsx', () => ({
  read: (...args: unknown[]) => mocks.xlsxRead(...args),
  utils: {
    sheet_to_json: (...args: unknown[]) => mocks.sheetToJson(...args),
  },
}));

vi.mock('@/infrastructure/api', () => ({
  workspaceAPI: {
    // Standalone viewers read through the owning workspace ID, never a bare path.
    readWorkspaceFile: (...args: unknown[]) => mocks.readFileContent(...args),
  },
}));

vi.mock('@/infrastructure/i18n', () => {
  const t = (key: string, options?: Record<string, string | number>) => {
    if (!options) return key;
    // The key stands in for the translation template; append values so tests can assert interpolation inputs.
    return `${key} ${Object.values(options).map(String).join(' ')}`;
  };
  const formatNumber = (value: number) => String(value);
  return {
    useI18n: () => ({ t, formatNumber }),
  };
});

vi.mock('@/shared/utils/logger', () => ({
  createLogger: () => ({
    error: vi.fn(),
    info: vi.fn(),
    warn: vi.fn(),
  }),
}));

vi.mock('@openbitfun/ui', () => ({
  Button: ({ children, icon: _icon, ...props }: React.ButtonHTMLAttributes<HTMLButtonElement> & { icon?: React.ReactNode }) => (
    <button {...props}>{children}</button>
  ),
  IconButton: ({ icon: _icon, ...props }: React.ButtonHTMLAttributes<HTMLButtonElement> & { icon?: React.ReactNode }) => (
    <button {...props} />
  ),
  Icon: ({ name }: { name: string }) => <span data-openbitfun-name={name} />,
  Select: ({
    onValueChange,
    options,
    size: _size,
    ...props
  }: Omit<React.SelectHTMLAttributes<HTMLSelectElement>, 'size'> & {
    onValueChange?: (value: string | number) => void;
    options: readonly Array<{ label: string; value: string | number }>;
    size?: 'sm' | 'md' | 'lg';
  }) => (
    <select
      {...props}
      onChange={event => {
        props.onChange?.(event);
        const option = options.find(candidate => String(candidate.value) === event.currentTarget.value);
        if (option) {
          onValueChange?.(option.value);
        }
      }}
    >
      {options.map(option => (
        <option key={option.value} value={option.value}>{option.label}</option>
      ))}
    </select>
  ),
  Toolbar: ({ leading, trailing, className }: { leading?: React.ReactNode; trailing?: React.ReactNode; className?: string }) => (
    <div className={className}>{leading}{trailing}</div>
  ),
  ToolbarGroup: ({ children }: React.PropsWithChildren) => <div>{children}</div>,
  ToolbarSeparator: () => <span />,
  Tooltip: ({ children }: React.PropsWithChildren) => <>{children}</>,
}));

async function flushAsyncWork(): Promise<void> {
  await act(async () => {
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    await new Promise(resolve => setTimeout(resolve, 0));
  });
}

describe('OfficeViewer', () => {
  let container: HTMLDivElement;
  let root: Root;

  beforeEach(() => {
    (globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
    container = document.createElement('div');
    document.body.appendChild(container);
    root = createRoot(container);
    mocks.readFileContent.mockReset();
    mocks.renderDocx.mockReset();
    mocks.xlsxRead.mockReset();
    mocks.sheetToJson.mockReset();
    mocks.readFileContent.mockResolvedValue('AQID');
    mocks.renderDocx.mockImplementation(async (_data: unknown, host: HTMLElement) => {
      const rendered = document.createElement('div');
      rendered.className = 'docx-mock-content';
      rendered.textContent = 'DOCX CONTENT';
      host.appendChild(rendered);
    });
    mocks.xlsxRead.mockReturnValue({
      SheetNames: ['Sheet1', 'Data'],
      Sheets: { Sheet1: { name: 'Sheet1' }, Data: { name: 'Data' } },
    });
    mocks.sheetToJson.mockReturnValue([['Name', 'Score'], ['Ada', '10'], ['Grace', '20']]);
  });

  afterEach(() => {
    act(() => root.unmount());
    container.remove();
    vi.clearAllMocks();
  });

  it('renders a docx through docx-preview after reading base64 content', async () => {
    act(() => root.render(<OfficeViewer workspaceId="workspace-remote" filePath="/remote/notes.docx" fileName="notes.docx" />));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(mocks.readFileContent).toHaveBeenCalledWith('workspace-remote', '/remote/notes.docx', 'base64');
    expect(mocks.renderDocx).toHaveBeenCalledOnce();
    const [data, host] = mocks.renderDocx.mock.calls[0] as [ArrayBuffer, HTMLElement];
    expect(data).toBeInstanceOf(ArrayBuffer);
    expect(new Uint8Array(data)).toEqual(new Uint8Array([1, 2, 3]));
    expect(host.className).toBe('openbitfun-office-viewer__document');
    expect(container.querySelector('[data-openbitfun-component="office-viewer"]')).not.toBeNull();
    expect(container.querySelector('.docx-mock-content')?.textContent).toBe('DOCX CONTENT');
    expect(container.textContent).not.toContain('editor.officeViewer.loading');
  });

  it('renders an xlsx workbook as a sheet table with a sheet selector', async () => {
    act(() => root.render(<OfficeViewer workspaceId="ws" filePath="/remote/budget.xlsx" fileName="budget.xlsx" />));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(mocks.xlsxRead).toHaveBeenCalledOnce();
    expect(mocks.renderDocx).not.toHaveBeenCalled();
    const sheetSelect = container.querySelector('.openbitfun-office-viewer__sheet-select') as HTMLSelectElement | null;
    expect(sheetSelect).not.toBeNull();
    expect([...(sheetSelect?.options ?? [])].map(option => option.value)).toEqual(['Sheet1', 'Data']);
    const table = container.querySelector('table.openbitfun-office-viewer__sheet');
    expect(table).not.toBeNull();
    expect(table?.querySelector('thead')?.textContent).toBe('NameScore');
    expect(table?.querySelectorAll('tbody tr')).toHaveLength(2);
  });

  it('shows a truncation notice when a sheet exceeds the row cap', async () => {
    mocks.sheetToJson.mockReturnValue(Array.from({ length: 1001 }, (_, index) => [`row-${index}`]));

    act(() => root.render(<OfficeViewer workspaceId="ws" filePath="/remote/big.xlsx" fileName="big.xlsx" />));
    await flushAsyncWork();
    await flushAsyncWork();

    // 1000 capped rows total, the first of which is the header row.
    expect(container.querySelectorAll('tbody tr')).toHaveLength(999);
    expect(container.textContent).toContain('editor.officeViewer.rowsTruncated');
  });

  it('degrades loudly with a download action for pptx instead of a broken preview', async () => {
    act(() => root.render(<OfficeViewer workspaceId="ws" filePath="/remote/slides.pptx" fileName="slides.pptx" />));
    await flushAsyncWork();
    await flushAsyncWork();

    expect(mocks.renderDocx).not.toHaveBeenCalled();
    expect(mocks.xlsxRead).not.toHaveBeenCalled();
    const unsupported = container.querySelector('[data-openbitfun-part="unsupported"]');
    expect(unsupported).not.toBeNull();
    expect(unsupported?.textContent).toContain('editor.officeViewer.unsupportedFormat');
    expect(unsupported?.textContent).toContain('pptx');
    expect(unsupported?.textContent).toContain('editor.officeViewer.download');
  });

  it('shows an error state with retry when the file cannot be read', async () => {
    mocks.readFileContent.mockRejectedValue(new Error('disk gone'));

    act(() => root.render(<OfficeViewer workspaceId="ws" filePath="/remote/missing.docx" fileName="missing.docx" isActiveTab={false} />));
    await flushAsyncWork();
    await flushAsyncWork();

    const error = container.querySelector('[data-openbitfun-part="error"]');
    expect(error).not.toBeNull();
    expect(error?.textContent).toContain('editor.officeViewer.loadFailedWithMessage');
    expect(error?.textContent).toContain('disk gone');
    expect(error?.textContent).toContain('editor.common.retry');
    expect(mocks.renderDocx).not.toHaveBeenCalled();
  });
});
