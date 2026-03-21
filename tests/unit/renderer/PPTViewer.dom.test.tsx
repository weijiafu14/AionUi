import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import React from 'react';

// Mock ipcBridge
vi.mock('@/common', () => ({
  ipcBridge: {
    document: {
      convert: { invoke: vi.fn() },
    },
    shell: {
      openFile: { invoke: vi.fn() },
    },
  },
}));

// Mock i18next with a stable t function to avoid infinite useEffect loops
// caused by a new function reference being created on every render
vi.mock('react-i18next', () => {
  const t = (k: string) => k;
  return { useTranslation: () => ({ t }) };
});

// Mock PDFViewer — mock both alias and relative resolution paths
vi.mock('@renderer/pages/conversation/Preview/components/viewers/PDFViewer', () => ({
  default: ({ filePath }: { filePath: string }) => <div data-testid='pdf-viewer'>{filePath}</div>,
}));

// Mock PreviewToolbarExtrasContext to avoid deep renderer imports
vi.mock('@renderer/pages/conversation/Preview/context/PreviewToolbarExtrasContext', () => ({
  usePreviewToolbarExtras: () => null,
}));

// Mock @icon-park/react to avoid large icon bundle
vi.mock('@icon-park/react', () => ({}));

// Mock @arco-design/web-react Button for jsdom compatibility
vi.mock('@arco-design/web-react', () => ({
  Button: ({
    children,
    onClick,
    disabled,
  }: {
    children: React.ReactNode;
    onClick?: () => void;
    disabled?: boolean;
  }) => (
    <button onClick={onClick} disabled={disabled}>
      {children}
    </button>
  ),
}));

import { ipcBridge } from '@/common';
import PPTViewer from '@renderer/pages/conversation/Preview/components/viewers/office/PPTViewer';

const mockSlides = [
  { slideNumber: 1, title: 'Slide One', texts: ['Body text'], images: [] },
  { slideNumber: 2, title: 'Slide Two', texts: [], images: [] },
];

describe('PPTViewer', () => {
  beforeEach(() => vi.clearAllMocks());

  it('renders PDFViewer when ppt-pdf succeeds', async () => {
    vi.mocked(ipcBridge.document.convert.invoke).mockResolvedValue({
      to: 'ppt-pdf',
      result: { success: true, data: '/tmp/test.pdf' },
    });

    render(<PPTViewer filePath='/test.pptx' />);

    await waitFor(() => expect(screen.getByTestId('pdf-viewer')).toBeTruthy());
    expect(screen.getByText('/tmp/test.pdf')).toBeTruthy();
  });

  it('renders fallback slides when ppt-pdf fails', async () => {
    vi.mocked(ipcBridge.document.convert.invoke)
      .mockResolvedValueOnce({
        to: 'ppt-pdf',
        result: { success: false, error: 'LIBRE_OFFICE_NOT_FOUND' },
      })
      .mockResolvedValueOnce({
        to: 'ppt-json',
        result: { success: true, data: { slides: mockSlides } },
      });

    render(<PPTViewer filePath='/test.pptx' />);

    await waitFor(() => expect(screen.getByText('Slide One')).toBeTruthy());
    // Shows fallback hint badge
    expect(screen.getByText('preview.ppt.fallbackHint')).toBeTruthy();
  });

  it('navigates slides with ArrowRight key', async () => {
    vi.mocked(ipcBridge.document.convert.invoke)
      .mockResolvedValueOnce({
        to: 'ppt-pdf',
        result: { success: false, error: 'LIBRE_OFFICE_NOT_FOUND' },
      })
      .mockResolvedValueOnce({
        to: 'ppt-json',
        result: { success: true, data: { slides: mockSlides } },
      });

    render(<PPTViewer filePath='/test.pptx' />);
    await waitFor(() => screen.getByText('Slide One'));

    fireEvent.keyDown(window, { key: 'ArrowRight' });
    await waitFor(() => expect(screen.getByText('Slide Two')).toBeTruthy());
  });
});
