/**
 * @license
 * Copyright 2025 AionUi (aionui.com)
 * SPDX-License-Identifier: Apache-2.0
 */

import { ipcBridge } from "@/common";
import React, { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";
import PDFViewer from "../PDFViewer";
import PPTPlaceholder from "./PPTPlaceholder";

interface PPTViewerProps {
  filePath?: string;
  hideToolbar?: boolean;
}

type ViewerState =
  | { status: "loading" }
  | { status: "pdf"; pdfPath: string }
  | { status: "error"; message: string };

const PPTViewer: React.FC<PPTViewerProps> = ({
  filePath,
  hideToolbar = false,
}) => {
  const { t } = useTranslation();
  const [state, setState] = useState<ViewerState>({ status: "loading" });

  useEffect(() => {
    if (!filePath) {
      setState({
        status: "error",
        message: t("preview.errors.missingFilePath"),
      });
      return;
    }

    const load = async () => {
      setState({ status: "loading" });
      try {
        const pdfResp = await ipcBridge.document.convert.invoke({
          filePath,
          to: "ppt-pdf",
        });
        if (
          pdfResp.to === "ppt-pdf" &&
          pdfResp.result.success &&
          pdfResp.result.data
        ) {
          setState({ status: "pdf", pdfPath: pdfResp.result.data });
          return;
        }
      } catch {
        // fall through to error
      }
      setState({ status: "error", message: t("preview.ppt.loadFailed") });
    };

    void load();
  }, [filePath, t]);

  if (state.status === "loading") {
    return (
      <div className="flex items-center justify-center h-full">
        <div className="text-14px text-t-secondary">
          {t("preview.ppt.loading")}
        </div>
      </div>
    );
  }

  if (state.status === "pdf") {
    return (
      <div className="h-full w-full flex flex-col">
        {!hideToolbar && (
          <div className="flex items-center justify-between h-40px px-12px bg-bg-2 shrink-0">
            <span className="text-11px text-t-tertiary">
              {t("preview.ppt.viaLibreOffice")}
            </span>
          </div>
        )}
        <div className="flex-1 overflow-hidden">
          <PDFViewer filePath={state.pdfPath} hideToolbar />
        </div>
      </div>
    );
  }

  return <PPTPlaceholder filePath={filePath} />;
};

export default PPTViewer;
