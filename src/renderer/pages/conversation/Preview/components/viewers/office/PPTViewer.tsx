/**
 * @license
 * Copyright 2025 AionUi (aionui.com)
 * SPDX-License-Identifier: Apache-2.0
 */

import { ipcBridge } from '@/common';
import type { PPTSlideData } from '@/common/types/conversion';
import { Button } from '@arco-design/web-react';
import React, { useCallback, useEffect, useState } from 'react';
import { useTranslation } from 'react-i18next';
import PDFViewer from '../PDFViewer';

interface PPTViewerProps {
  filePath?: string;
  hideToolbar?: boolean;
}

type ViewerState =
  | { status: 'loading' }
  | { status: 'pdf'; pdfPath: string }
  | { status: 'fallback-loading' }
  | {
      status: 'slides';
      slides: PPTSlideData[];
      currentIndex: number;
      usedFallback: true;
    }
  | { status: 'error'; message: string };

const PPTViewer: React.FC<PPTViewerProps> = ({ filePath, hideToolbar = false }) => {
  const { t } = useTranslation();
  const [state, setState] = useState<ViewerState>({ status: 'loading' });

  // Load: try ppt-pdf, fall back to ppt-json on any failure
  useEffect(() => {
    if (!filePath) {
      setState({
        status: 'error',
        message: t('preview.errors.missingFilePath'),
      });
      return;
    }

    const load = async () => {
      setState({ status: 'loading' });

      // Step 1: try LibreOffice PDF conversion
      try {
        const pdfResp = await ipcBridge.document.convert.invoke({
          filePath,
          to: 'ppt-pdf',
        });
        if (pdfResp.to === 'ppt-pdf' && pdfResp.result.success && pdfResp.result.data) {
          setState({ status: 'pdf', pdfPath: pdfResp.result.data });
          return;
        }
      } catch {
        // proceed to fallback
      }

      // Step 2: fallback — pptx2json
      setState({ status: 'fallback-loading' });
      try {
        const jsonResp = await ipcBridge.document.convert.invoke({
          filePath,
          to: 'ppt-json',
        });
        if (jsonResp.to === 'ppt-json' && jsonResp.result.success && jsonResp.result.data) {
          setState({
            status: 'slides',
            slides: jsonResp.result.data.slides,
            currentIndex: 0,
            usedFallback: true,
          });
          return;
        }
        throw new Error(jsonResp.to === 'ppt-json' ? (jsonResp.result.error ?? '') : '');
      } catch (err) {
        setState({
          status: 'error',
          message: err instanceof Error ? err.message : t('preview.ppt.loadFailed'),
        });
      }
    };

    void load();
  }, [filePath, t]);

  // Keyboard navigation
  useEffect(() => {
    if (state.status !== 'slides') return;
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'ArrowRight') {
        setState((prev) =>
          prev.status === 'slides' && prev.currentIndex < prev.slides.length - 1
            ? { ...prev, currentIndex: prev.currentIndex + 1 }
            : prev
        );
      } else if (e.key === 'ArrowLeft') {
        setState((prev) =>
          prev.status === 'slides' && prev.currentIndex > 0 ? { ...prev, currentIndex: prev.currentIndex - 1 } : prev
        );
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [state.status]);

  const handleOpenInSystem = useCallback(async () => {
    if (!filePath) return;
    try {
      await ipcBridge.shell.openFile.invoke(filePath);
    } catch {
      // ignore
    }
  }, [filePath]);

  // --- Render ---

  if (state.status === 'loading' || state.status === 'fallback-loading') {
    const label = state.status === 'loading' ? t('preview.ppt.loading') : t('preview.ppt.converting');
    return (
      <div className='flex items-center justify-center h-full'>
        <div className='text-14px text-t-secondary'>{label}</div>
      </div>
    );
  }

  if (state.status === 'pdf') {
    return (
      <div className='h-full w-full flex flex-col'>
        {!hideToolbar && (
          <div className='flex items-center justify-between h-40px px-12px bg-bg-2 shrink-0'>
            <span className='text-11px text-t-tertiary'>{t('preview.ppt.viaLibreOffice')}</span>
          </div>
        )}
        <div className='flex-1 overflow-hidden'>
          <PDFViewer filePath={state.pdfPath} hideToolbar />
        </div>
      </div>
    );
  }

  if (state.status === 'error') {
    return (
      <div className='flex items-center justify-center h-full'>
        <div className='text-center'>
          <div className='text-16px text-t-error mb-8px'>{state.message}</div>
          {filePath && (
            <Button size='small' onClick={handleOpenInSystem}>
              {t('preview.openInSystemApp')}
            </Button>
          )}
        </div>
      </div>
    );
  }

  // state.status === 'slides'
  const { slides, currentIndex } = state;
  const slide = slides[currentIndex];
  const total = slides.length;

  return (
    <div className='h-full w-full flex flex-col bg-bg-1'>
      {/* Toolbar */}
      {!hideToolbar && (
        <div className='flex items-center justify-between h-40px px-12px bg-bg-2 shrink-0'>
          <span className='text-11px text-t-tertiary'>{t('preview.ppt.fallbackHint')}</span>
          {filePath && (
            <Button size='mini' type='text' onClick={handleOpenInSystem}>
              {t('preview.openInSystemApp')}
            </Button>
          )}
        </div>
      )}

      {/* Slide content — fixed 16:9 aspect ratio card */}
      <div className='flex-1 overflow-hidden flex items-center justify-center p-16px'>
        <div
          className='w-full bg-bg-2 rd-8px shadow-sm overflow-hidden flex flex-col'
          style={{ aspectRatio: '16/9', maxHeight: '100%' }}
        >
          <div className='flex-1 overflow-y-auto p-24px'>
            {slide.title && (
              <div className='text-18px font-semibold text-t-primary mb-12px leading-snug'>{slide.title}</div>
            )}
            {slide.texts.map((text, i) => (
              <p key={i} className='text-14px text-t-secondary leading-relaxed mb-6px'>
                {text}
              </p>
            ))}
            {slide.images.length > 0 && (
              <div className='flex flex-wrap gap-8px mt-12px'>
                {slide.images.map((src, i) => (
                  <img
                    key={i}
                    src={src}
                    alt=''
                    className='object-contain rd-4px'
                    style={{ maxHeight: '200px', maxWidth: '48%' }}
                  />
                ))}
              </div>
            )}
            {!slide.title && slide.texts.length === 0 && slide.images.length === 0 && (
              <div className='text-13px text-t-tertiary'>{t('preview.ppt.noContent')}</div>
            )}
          </div>
        </div>
      </div>

      {/* Navigation footer */}
      <div className='flex items-center justify-center gap-16px h-48px shrink-0 border-t border-border-base'>
        <Button
          size='small'
          disabled={currentIndex === 0}
          onClick={() =>
            setState((prev) => (prev.status === 'slides' ? { ...prev, currentIndex: prev.currentIndex - 1 } : prev))
          }
        >
          {t('preview.ppt.prevSlide')}
        </Button>
        <span className='text-13px text-t-secondary min-w-48px text-center'>
          {t('preview.ppt.slideCount', { current: currentIndex + 1, total })}
        </span>
        <Button
          size='small'
          disabled={currentIndex === total - 1}
          onClick={() =>
            setState((prev) => (prev.status === 'slides' ? { ...prev, currentIndex: prev.currentIndex + 1 } : prev))
          }
        >
          {t('preview.ppt.nextSlide')}
        </Button>
      </div>
    </div>
  );
};

export default PPTViewer;
