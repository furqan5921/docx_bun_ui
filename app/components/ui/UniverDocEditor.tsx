// Import styles for preset mode
import { DocSelectionManagerService } from "@univerjs/docs";
import { IRenderManagerService } from "@univerjs/engine-render";
import "@univerjs/preset-docs-core/lib/index.css";
import { Download, Upload } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import { Button } from "~/components/ui/button";
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetHeader,
  SheetTitle,
} from "~/components/ui/sheet";
import { CursorNavigationPlugin } from "~/plugins/CursorNavigationPlugin";
import { CustomQuickInsertPlugin } from "~/plugins/CustomQuickInsertPlugin";
import { HorizontalLineSpacingPlugin } from "~/plugins/HorizontalLineSpacingPlugin";
import { convertDocxToUniverData } from "~/utils/docx-converter";
import CustomQuickInsertMenu from "../CustomQuickInsertMenu";


interface UniverDocEditorProps {
  initialFile?: File;
}

export function UniverDocEditor({ initialFile }: UniverDocEditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [importing, setImporting] = useState(false);
  // biome-ignore lint/suspicious/noExplicitAny: IDocumentData type from converter
  const [documentData, setDocumentData] = useState<any>(null);
  const [fileName, setFileName] = useState<string>("");
  const timeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const lifecycleDisposeRef = useRef<{ dispose: () => void } | null>(null);
  // biome-ignore lint/suspicious/noExplicitAny: Univer API type is complex
  const univerAPIRef = useRef<any>(null);
  // biome-ignore lint/suspicious/noExplicitAny: Univer instance type is complex
  const univerInstanceRef = useRef<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Quick insert menu state
  const [menuVisible, setMenuVisible] = useState(false);
  const [menuPosition, setMenuPosition] = useState<
    { x: number; y: number } | undefined
  >();
  const quickInsertServiceRef = useRef<any>(null);

  // AI Assistant Sheet state
  const [showAISheet, setShowAISheet] = useState(false);
  const [selectedText, setSelectedText] = useState("");
  const [aiPrompt, setAiPrompt] = useState("");
  const [selectionRange, setSelectionRange] = useState<{startOffset: number, endOffset: number} | null>(null);

  // Load initial file if provided
  useEffect(() => {
    if (initialFile) {
      console.log(`📂 Loading initial file: ${initialFile.name}`);
      setFileName(initialFile.name);
      setImporting(true);
      setLoading(true);

      initialFile.arrayBuffer().then(async (arrayBuffer) => {
        try {
          const convertedData = await convertDocxToUniverData(arrayBuffer);
          console.log("✅ Initial file converted successfully");
          setDocumentData(convertedData);
        } catch (err) {
          console.error("❌ Failed to import initial file:", err);
          setError(
            err instanceof Error ? err.message : "Failed to import document",
          );
          setLoading(false);
          setImporting(false);
        }
      });
    }
  }, [initialFile]);

  // biome-ignore lint/correctness/useExhaustiveDependencies: documentData accessed in nested function
  useEffect(() => {
    if (!containerRef.current) return;

    // Dynamically import Univer packages using PRESET MODE
    const initUniver = async () => {
      try {
        const [presetsModule, presetDocsCore, presetLocale] = await Promise.all(
          [
            import("@univerjs/presets"),
            import("@univerjs/preset-docs-core"),
            import("@univerjs/preset-docs-core/locales/en-US"),
          ],
        );

        const { createUniver, LocaleType } = presetsModule;
        const { UniverDocsCorePreset } = presetDocsCore;
        const UniverPresetDocsCoreEnUS = presetLocale.default || presetLocale;

        if (!containerRef.current) {
          throw new Error("Container ref is not available");
        }

        // Initialize Univer using PRESET MODE
        const { univerAPI, univer } = createUniver({
          locale: LocaleType.EN_US,
          locales: {
            [LocaleType.EN_US]: UniverPresetDocsCoreEnUS,
          },
          presets: [
            UniverDocsCorePreset({
              container: containerRef.current,
            }),
          ],
        });

        // Register the CustomQuickInsertPlugin after Univer is created
        univer.registerPlugin(CustomQuickInsertPlugin);
        console.log("✅ CustomQuickInsertPlugin registered successfully");

        // Register the CursorNavigationPlugin to fix word/line navigation shortcuts
        // This enables: Ctrl+Arrow (word nav), Home/End (line nav) on Windows
        //              Option+Arrow (word nav), Cmd+Arrow (line nav) on macOS
        univer.registerPlugin(CursorNavigationPlugin);
        console.log("✅ CursorNavigationPlugin registered successfully");

        // Register the HorizontalLineSpacingPlugin for improved horizontal line spacing
        univer.registerPlugin(HorizontalLineSpacingPlugin);
        console.log("✅ HorizontalLineSpacingPlugin registered successfully");

        univerAPIRef.current = univerAPI;
        univerInstanceRef.current = univer;

        // Setup quick insert menu event listeners
        const handleQuickInsertShow = (event: CustomEvent) => {
          console.log(
            "[UniverDocEditor] Quick insert show event:",
            event.detail,
          );

          // Get service if not already available
          if (!quickInsertServiceRef.current) {
            quickInsertServiceRef.current = (
              window as any
            ).__customQuickInsertService;
          }

          const { bounds } = event.detail;
          if (bounds) {
            const univerContainer = containerRef.current;

            if (univerContainer) {
              const canvasElement = univerContainer.querySelector(
                "canvas"
              ) as HTMLCanvasElement | null;

              if (canvasElement) {
                const canvasRect = canvasElement.getBoundingClientRect();

                // Find ANY scrollable parent element
                let scrollTop = 0;
                let scrollLeft = 0;
                let scrollableElement: HTMLElement | null = null;

                // Check all parent elements for scroll
                let element: HTMLElement | null = canvasElement.parentElement;
                while (element && element !== document.body) {
                  const hasVerticalScroll = element.scrollHeight > element.clientHeight;
                  const hasHorizontalScroll = element.scrollWidth > element.clientWidth;
                  const computedStyle = window.getComputedStyle(element);
                  const overflowY = computedStyle.overflowY;
                  const overflowX = computedStyle.overflowX;

                  if (
                    (hasVerticalScroll && (overflowY === 'auto' || overflowY === 'scroll')) ||
                    (hasHorizontalScroll && (overflowX === 'auto' || overflowX === 'scroll'))
                  ) {
                    scrollableElement = element;
                    scrollTop = element.scrollTop;
                    scrollLeft = element.scrollLeft;
                    console.log("[UniverDocEditor] Found scrollable container:", {
                      element: element.className,
                      scrollTop,
                      scrollLeft,
                      scrollHeight: element.scrollHeight,
                      clientHeight: element.clientHeight
                    });
                    break;
                  }
                  element = element.parentElement;
                }

                console.log("[UniverDocEditor] Canvas rect:", canvasRect);
                console.log("[UniverDocEditor] Raw bounds:", bounds);
                console.log("[UniverDocEditor] Final scroll offset:", { scrollTop, scrollLeft });

                // Calculate viewport coordinates (bounds are in document space, convert to viewport)
                // bounds.bottom is absolute document position, subtract scroll to get viewport position
                const viewportX = bounds.left + canvasRect.left - scrollLeft;
                const viewportY = bounds.bottom + canvasRect.top - scrollTop + 5;

                console.log(
                  "[UniverDocEditor] Viewport position:",
                  viewportX,
                  viewportY
                );

                setMenuPosition({ x: viewportX, y: viewportY });
                setMenuVisible(true);
              } else {
                console.warn("[UniverDocEditor] Canvas element not found");
                setMenuPosition({ x: bounds.left, y: bounds.bottom + 5 });
                setMenuVisible(true);
              }
            }
          }
        };

        const handleQuickInsertClose = () => {
          console.log("[UniverDocEditor] Quick insert close event");
          setMenuVisible(false);
        };

        window.addEventListener(
          "univer:quick-insert-show",
          handleQuickInsertShow as EventListener,
        );
        window.addEventListener(
          "univer:quick-insert-close",
          handleQuickInsertClose as EventListener,
        );

        const createDocument = () => {
          try {
            setError(null);

            if (documentData) {
              console.log("📄 Creating document with imported data");
              console.log(
                "🔍 Document data structure:",
                JSON.stringify(
                  {
                    hasBody: !!documentData.body,
                    dataStreamLength:
                      documentData.body?.dataStream?.length || 0,
                    paragraphsCount: documentData.body?.paragraphs?.length || 0,
                    textRunsCount: documentData.body?.textRuns?.length || 0,
                    bodyTablesCount: documentData.body?.tables?.length || 0,
                    hasTableSource: !!documentData.tableSource,
                    tableSourceKeys: documentData.tableSource ? Object.keys(documentData.tableSource) : [],
                    dataStreamPreview:
                      documentData.body?.dataStream?.substring(0, 100) ||
                      "empty",
                    firstParagraph: documentData.body?.paragraphs?.[0] || null,
                    firstTextRun: documentData.body?.textRuns?.[0] || null,
                  },
                  null,
                  2,
                ),
              );

              // Debug: Log table control characters in dataStream
              if (documentData.body?.tables?.length > 0) {
                const ds = documentData.body.dataStream;
                console.log("🔍 Table control characters in dataStream:");
                for (let i = 0; i < ds.length; i++) {
                  const code = ds.charCodeAt(i);
                  if (code >= 0x0E && code <= 0x1F) {
                    console.log(`   Position ${i}: char code ${code} (0x${code.toString(16)})`);
                  }
                }
                console.log("🔍 Body tables:", documentData.body.tables);
                console.log("🔍 TableSource:", documentData.tableSource);

                // Verify each table's startIndex points to \x1A (26) and endIndex-1 to \x0F (15)
                documentData.body.tables.forEach((table: any, idx: number) => {
                  const startChar = ds.charCodeAt(table.startIndex);
                  const endChar = ds.charCodeAt(table.endIndex - 1);
                  const isValid = startChar === 0x1A && endChar === 0x0F;
                  console.log(`📋 Table ${idx} (${table.tableId}):`);
                  console.log(`   startIndex=${table.startIndex}, char=0x${startChar.toString(16)} (${startChar === 0x1A ? '✅ TABLE_START' : '❌ WRONG'})`);
                  console.log(`   endIndex=${table.endIndex}, char at endIndex-1=0x${endChar.toString(16)} (${endChar === 0x0F ? '✅ TABLE_END' : '❌ WRONG'})`);
                  console.log(`   Valid: ${isValid ? '✅ YES' : '❌ NO'}`);

                  // Check if tableSource has this table
                  const hasTableSource = documentData.tableSource && documentData.tableSource[table.tableId];
                  console.log(`   TableSource entry exists: ${hasTableSource ? '✅ YES' : '❌ NO'}`);
                  if (hasTableSource) {
                    const ts = documentData.tableSource[table.tableId];
                    console.log(`   TableSource details: ${ts.tableRows?.length} rows, ${ts.tableColumns?.length} columns`);
                  }
                });
              }

              // Ensure the document data has valid content
              if (
                !documentData.body ||
                !documentData.body.dataStream ||
                documentData.body.dataStream.length === 0
              ) {
                console.warn(
                  "⚠️ Document data appears empty, creating with sample content",
                );
                univerAPI.createUniverDoc({
                  body: {
                    dataStream: "Sample Document Content\r\n",
                    paragraphs: [
                      { startIndex: 0, paragraphStyle: { horizontalAlign: 0 } },
                      {
                        startIndex: 24,
                        paragraphStyle: { horizontalAlign: 0 },
                      },
                    ],
                    textRuns: [{ st: 0, ed: 23, ts: {} }],
                    sectionBreaks: [{ startIndex: 24 }],
                  },
                });
              } else {
                // CRITICAL: Pass complete document data including tableSource
                const { body, documentStyle, tableSource, lists } = documentData;

                // Build the document payload matching IDocumentData structure
                const docPayload: any = {
                  body: {
                    dataStream: body.dataStream || "\r\n",
                    textRuns: body.textRuns || [],
                    paragraphs: body.paragraphs || [],
                    sectionBreaks: body.sectionBreaks || [{ startIndex: (body.dataStream?.length || 1) - 1 }],
                    tables: body.tables || [],
                    customBlocks: body.customBlocks || [],
                    customRanges: body.customRanges || [],
                  },
                  documentStyle: documentStyle || {
                    pageSize: { width: 595.27, height: 841.89 },
                    documentFlavor: 1,
                    marginTop: 72,
                    marginBottom: 72,
                    marginLeft: 72,
                    marginRight: 72,
                    renderConfig: {
                      zeroWidthParagraphBreak: 0,
                      vertexAngle: 0,
                      centerAngle: 0,
                      background: { rgb: "#ffffff" },
                    },
                  },
                  // CRITICAL: tableSource must be top-level (from IReferenceSource interface)
                  tableSource: tableSource || {},
                };

                // Include lists if present
                if (lists && Object.keys(lists).length > 0) {
                  docPayload.lists = lists;
                }

                console.log("📋 Creating doc with tableSource keys:", Object.keys(docPayload.tableSource || {}));
                const doc = univerAPI.createUniverDoc(docPayload);

                // CRITICAL WORKAROUND: Univer's _buildTableCache() is called during createUniverDoc()
                // when getSnapshot().tableSource might still be null due to timing issues.
                // We need to patch the data model and trigger a view model reset.
                if (tableSource && Object.keys(tableSource).length > 0 && body.tables?.length > 0) {
                  setTimeout(async () => {
                    try {
                      const activeDoc = univerAPI.getActiveDocument();
                      if (activeDoc && activeDoc.getId() === doc.getId()) {
                        // Access internal document data model
                        const docAny = activeDoc as unknown as {
                          _injector?: { get: (token: unknown) => unknown };
                          _documentDataModel?: {
                            getUnitId: () => string;
                            getSnapshot: () => { tableSource?: Record<string, unknown>; body?: { tables?: unknown[] } };
                          };
                        };

                        if (docAny._injector && docAny._documentDataModel) {
                          const dataModel = docAny._documentDataModel;
                          const modelSnapshot = dataModel.getSnapshot();

                          // Patch tableSource onto the data model's snapshot
                          if (!modelSnapshot.tableSource) {
                            (modelSnapshot as Record<string, unknown>).tableSource = tableSource;
                            console.log("🔧 Patched tableSource onto data model snapshot");
                          }
                          if (modelSnapshot?.body && !Array.isArray(modelSnapshot.body.tables)) {
                            (modelSnapshot.body as Record<string, unknown>).tables = body.tables;
                            console.log("🔧 Patched body.tables onto data model snapshot");
                          }

                          if (modelSnapshot?.tableSource) {
                            const injector = docAny._injector;
                            const unitId = dataModel.getUnitId();

                            // Get render manager and trigger view model reset
                            const renderManager = injector.get(IRenderManagerService) as {
                              getRenderById: (id: string) => {
                                with: (tok: unknown) => {
                                  get: (tok: unknown) => unknown;
                                };
                              } | null;
                            };

                            if (renderManager) {
                              const renderUnit = renderManager.getRenderById(unitId);
                              if (renderUnit) {
                                // Get the view model token
                                const DocViewModelManagerServiceToken = Symbol.for("DocViewModelManagerService");
                                const viewModelManager = renderUnit.with(DocViewModelManagerServiceToken).get(DocViewModelManagerServiceToken) as {
                                  getViewModel: () => {
                                    reset: () => void;
                                  } | null;
                                } | null;

                                if (viewModelManager) {
                                  const viewModel = viewModelManager.getViewModel();
                                  if (viewModel && typeof viewModel.reset === "function") {
                                    console.log("🔄 Triggering view model reset to rebuild table cache...");
                                    viewModel.reset();

                                    // Also trigger skeleton calculation
                                    const DocSkeletonManagerServiceToken = Symbol.for("DocSkeletonManagerService");
                                    const skeletonManager = renderUnit.with(DocSkeletonManagerServiceToken).get(DocSkeletonManagerServiceToken) as {
                                      getCurrent: () => { calculate: () => void } | null;
                                    } | null;

                                    if (skeletonManager) {
                                      const skeleton = skeletonManager.getCurrent();
                                      if (skeleton && typeof skeleton.calculate === "function") {
                                        console.log("🔄 Triggering skeleton recalculation...");
                                        skeleton.calculate();
                                      }
                                    }
                                    console.log("✅ Table cache rebuild triggered");
                                  }
                                }
                              }
                            }
                          }
                        }
                      }
                    } catch (e) {
                      console.warn("⚠️ Could not trigger view model reset for table cache:", e);
                    }
                  }, 50);
                }
              }
            } else {
              console.log("📄 Creating blank document");
              // Create blank document with TRADITIONAL mode (page view with boundaries)
              univerAPI.createUniverDoc({
                body: {
                  dataStream: "\r\n",
                  paragraphs: [{ startIndex: 0, paragraphStyle: {} }],
                  sectionBreaks: [{ startIndex: 1 }],
                  tables: [],
                  customBlocks: [],
                },
                documentStyle: {
                  pageSize: { width: 595.27, height: 841.89 }, // A4 size
                  documentFlavor: 1, // TRADITIONAL - enables page view with boundaries
                  marginTop: 72,
                  marginBottom: 72,
                  marginLeft: 90,
                  marginRight: 90,
                  renderConfig: {
                    vertexAngle: 0,
                    centerAngle: 0,
                    background: { rgb: "#FFFFFF" },
                  },
                },
                tableSource: {},
              });
            }

            setTimeout(() => {
              const activeDoc = univerAPI.getActiveDocument();
              if (activeDoc) {
                console.log("✓ Document created:", activeDoc.getId());
                const snapshot = activeDoc.getSnapshot();
                const snapshotAny = snapshot as any;
                console.log(
                  "📋 Document snapshot:",
                  JSON.stringify(
                    {
                      hasBody: !!snapshot?.body,
                      dataStreamLength: snapshot?.body?.dataStream?.length || 0,
                      paragraphsCount: snapshot?.body?.paragraphs?.length || 0,
                      bodyTablesCount: snapshot?.body?.tables?.length || 0,
                      hasTableSource: !!snapshotAny?.tableSource,
                      tableSourceKeys: snapshotAny?.tableSource ? Object.keys(snapshotAny.tableSource) : [],
                    },
                    null,
                    2,
                  ),
                );
                // Log tableSource details if present
                if (snapshotAny?.tableSource && Object.keys(snapshotAny.tableSource).length > 0) {
                  console.log("✅ Snapshot has tableSource:", snapshotAny.tableSource);
                } else {
                  console.log("❌ Snapshot tableSource is EMPTY - tables won't render!");
                }
              }
            }, 200);

            setLoading(false);
            setImporting(false);
          } catch (e) {
            console.error("Failed to create document:", e);
            setError("Failed to create document");
            setLoading(false);
            setImporting(false);
          }
        };

        let contentCreated = false;

        const createContentOnce = () => {
          if (!contentCreated) {
            contentCreated = true;
            createDocument();
          }
        };

        // Force-apply alignment after document creation using Univer commands
        // This ensures alignment triggers proper layout recalculation and render invalidation
        const applyAlignmentFix = async () => {
          const api = univerAPIRef.current;
          if (!api) return;

          const doc = api.getActiveDocument();
          if (!doc) return;

          const snapshot = doc.getSnapshot();
          const paragraphs = snapshot.body?.paragraphs ?? [];

          console.log(`🔧 Applying alignment fix to ${paragraphs.length} paragraphs`);

          try {
            // Alignment is already applied during import in docx-converter.ts
            // No need to re-apply alignment via SetParagraphAlignCommand
            // (SetParagraphAlignCommand doesn't exist in @univerjs/docs v0.15.x)

            paragraphs.forEach((p: any, index: number) => {
              const alignment = p.paragraphStyle?.horizontalAlign;
              if (alignment !== undefined && alignment !== 0) {
                // Alignment is already in the data from import
                console.log(`  ✓ Paragraph ${index} at ${p.startIndex}: has alignment ${alignment}`);
              }
            });

            console.log("✅ Alignment check completed");
          } catch (err) {
            console.error("❌ Error checking alignment:", err);
          }
        };

        const disposable = univerAPI.addEvent(
          univerAPI.Event.LifeCycleChanged,
          // biome-ignore lint/suspicious/noExplicitAny: Univer lifecycle event structure
          ({ stage }: any) => {
            console.log("Univer lifecycle stage:", stage);

            if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
              createContentOnce();

              // Apply alignment fix after document is rendered
              setTimeout(() => {
                applyAlignmentFix();
              }, 100);

              setTimeout(() => {
                const activeDoc = univerAPI.getActiveDocument();
                if (activeDoc) {
                  console.log(
                    "Active document at Rendered:",
                    activeDoc.getId(),
                  );
                }

                // Focus the editor
                const focusEditor = () => {
                  const selectors = [
                    '[contenteditable="true"]',
                    ".univer-editor",
                    '[role="textbox"]',
                  ];

                  for (const selector of selectors) {
                    const editorElement = containerRef.current?.querySelector(
                      selector,
                    ) as HTMLElement;
                    if (editorElement) {
                      editorElement.focus();
                      editorElement.click();
                      console.log("Editor focused using selector:", selector);
                      return true;
                    }
                  }
                  return false;
                };

                if (!focusEditor()) {
                  setTimeout(() => {
                    focusEditor();
                  }, 200);
                }
              }, 300);
            }

            if (stage === univerAPI.Enum.LifecycleStages.Steady) {
              if (disposable) {
                disposable.dispose();
                lifecycleDisposeRef.current = null;
              }
            }
          },
        );
        lifecycleDisposeRef.current = disposable || null;

        timeoutRef.current = setTimeout(() => {
          if (!contentCreated) {
            console.log("Creating document via fallback timeout");
            createContentOnce();
          }
          if (lifecycleDisposeRef.current) {
            lifecycleDisposeRef.current.dispose();
            lifecycleDisposeRef.current = null;
          }
        }, 1500);
      } catch (e) {
        console.error("Failed to initialize Univer:", e);
        setError("Failed to initialize editor");
        setLoading(false);
      }
    };

    initUniver();

    // Listen for AI action events
    const handleAIAction = (event: Event) => {
      const customEvent = event as CustomEvent;
      console.log("[AI Event] Received:", customEvent.detail);
      if (customEvent.detail?.selectedText) {
        setSelectedText(customEvent.detail.selectedText);
        setSelectionRange(customEvent.detail.selectionRange || null);
        setShowAISheet(true);
      }
    };
    window.addEventListener("univer:ai-action", handleAIAction);

    return () => {
      if (timeoutRef.current) {
        clearTimeout(timeoutRef.current);
      }
      if (lifecycleDisposeRef.current) {
        lifecycleDisposeRef.current.dispose();
        lifecycleDisposeRef.current = null;
      }
      if (univerInstanceRef.current) {
        univerInstanceRef.current.dispose();
        univerInstanceRef.current = null;
        univerAPIRef.current = null;
      }
      // Cleanup event listeners
      window.removeEventListener("univer:quick-insert-show", () => { });
      window.removeEventListener("univer:quick-insert-close", () => { });
      window.removeEventListener("univer:ai-action", () => { });
    };
  }, [documentData]);

  // Handle DOCX file import
  const handleFileImport = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.name.endsWith(".docx") && !file.name.endsWith(".doc")) {
      setError("Please select a valid DOCX file");
      return;
    }

    setImporting(true);
    setLoading(true);
    setError(null);
    setFileName(file.name);

    try {
      console.log(`📂 Importing DOCX file: ${file.name}`);

      const arrayBuffer = await file.arrayBuffer();
      const convertedData = await convertDocxToUniverData(arrayBuffer);

      console.log("✅ DOCX converted successfully");
      setDocumentData(convertedData);
    } catch (err) {
      console.error("❌ Failed to import DOCX:", err);
      setError(
        err instanceof Error ? err.message : "Failed to import document",
      );
      setLoading(false);
      setImporting(false);
    } finally {
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const triggerFileImport = () => {
    fileInputRef.current?.click();
  };

  const handleExport = async () => {
    try {
      console.log("[DOCX Export] Starting SERVER-SIDE export...");

      if (!univerAPIRef.current) {
        setError("Editor not initialized");
        return;
      }

      const activeDoc = univerAPIRef.current.getActiveDocument();
      if (!activeDoc) {
        setError("No active document");
        return;
      }

      // CRITICAL FIX: Force document to commit any pending edits before snapshot
      // This ensures getSnapshot() includes the very latest user changes
      console.log("[DOCX Export] Flushing pending edits...");

      // Trigger a blur event to force Univer to commit any buffered text
      const editorElement = document.querySelector('.univer-render-canvas');
      if (editorElement) {
        (editorElement as HTMLElement).blur();
        // Small delay to let Univer process the blur and commit changes
        await new Promise(resolve => setTimeout(resolve, 100));
      }

      const snapshot = activeDoc.getSnapshot();

      // Log full IDocumentData contents
      console.log("[DOCX Export] ========== FULL IDocumentData Contents ==========");
      console.log("[DOCX Export] Complete snapshot object:", snapshot);
      console.log("[DOCX Export] Pretty printed:", JSON.stringify(snapshot, null, 2));

      console.log("[DOCX Export] Quick summary:", {
        textRuns: snapshot.body?.textRuns?.length || 0,
        dataStreamLength: snapshot.body?.dataStream?.length || 0,
        paragraphs: snapshot.body?.paragraphs?.length || 0,
        dataStreamPreview: snapshot.body?.dataStream?.substring(0, 200) || "empty",
        dataStreamEndPreview: snapshot.body?.dataStream?.substring(Math.max(0, (snapshot.body?.dataStream?.length || 0) - 50)) || "empty",
      });
      console.log("[DOCX Export] ===================================================");

      // Prepare filename
      const exportFileName = fileName
        ? `${fileName.replace(/\.[^/.]+$/, "")}_edited.docx`
        : "document_edited.docx";

      console.log("[DOCX Export] Sending to server for processing...");

      // Send document data to server for DOCX generation
      const response = await fetch("/api/export", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          documentData: snapshot,
          filename: exportFileName,
        }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(
          errorData.message ||
          `Server error: ${response.status} ${response.statusText}`
        );
      }

      console.log("[DOCX Export] Receiving file from server...");

      // Get the blob from server response
      const blob = await response.blob();

      console.log("[DOCX Export] File received, size:", blob.size, "bytes");

      // Download using file-saver
      const FileSaver = await import("file-saver");
      FileSaver.saveAs(blob, exportFileName);

      console.log(
        "[DOCX Export] ✓ Document exported successfully via SERVER as",
        exportFileName
      );
    } catch (err) {
      console.error("[DOCX Export] Failed:", err);
      setError(
        "Failed to export document: " +
        (err instanceof Error ? err.message : String(err))
      );
    }
  };

  const handleMenuSelect = (menu: any) => {
    console.log("[UniverDocEditor] Menu selected:", menu);

    // Get service if not already available
    const service =
      quickInsertServiceRef.current ||
      (window as any).__customQuickInsertService;

    if (service) {
      quickInsertServiceRef.current = service;
      service.emitMenuSelected(menu);
    } else {
      console.warn(
        "[UniverDocEditor] Service not available for menu selection",
      );
    }

    setMenuVisible(false);
  };

  const handleMenuClose = () => {
    console.log("[UniverDocEditor] Menu closed by user");

    // Get service if not already available
    const service =
      quickInsertServiceRef.current ||
      (window as any).__customQuickInsertService;

    if (service) {
      quickInsertServiceRef.current = service;
      service.closePopup();
    }

    setMenuVisible(false);
  };

  return (
    <div className="relative flex h-full flex-col bg-gray-100">
      <input
        ref={fileInputRef}
        type="file"
        accept=".docx,.doc"
        onChange={handleFileImport}
        className="hidden"
      />

      {/* Quick Insert Menu */}
      {menuVisible && (
        <CustomQuickInsertMenu
          service={
            quickInsertServiceRef.current ||
            (window as any).__customQuickInsertService
          }
          visible={menuVisible}
          position={menuPosition}
          onSelect={handleMenuSelect}
          onClose={handleMenuClose}
        />
      )}

      {/* Header Toolbar - Production Ready UI */}
      {!error && (
        <div className="flex-shrink-0 border-b border-gray-200 bg-white px-4 py-2 shadow-sm">
          <div className="flex items-center justify-between">
            {/* Left side - Document info */}
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-2">
                <svg className="h-5 w-5 text-blue-600" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zm4 18H6V4h7v5h5v11z"/>
                </svg>
                <span className="font-medium text-gray-700 text-sm">
                  {fileName || "Untitled Document"}
                </span>
              </div>
              {importing && (
                <span className="text-xs text-blue-600 animate-pulse">
                  Importing...
                </span>
              )}
            </div>

            {/* Right side - Action buttons */}
            <div className="flex items-center gap-2">
              <Button
                variant="outline"
                size="sm"
                onClick={triggerFileImport}
                disabled={importing || loading}
                className="border-gray-300 hover:bg-gray-50 hover:border-gray-400"
              >
                <Upload className="h-4 w-4 mr-2" />
                Import
              </Button>
              <Button
                variant="outline"
                size="sm"
                onClick={handleExport}
                disabled={loading}
                className="border-gray-300 hover:bg-gray-50 hover:border-gray-400"
              >
                <Download className="h-4 w-4 mr-2" />
                Export
              </Button>
              <div className="w-px h-6 bg-gray-300 mx-1" />
              <Button
                size="sm"
                onClick={async () => {
                  console.log("[AI Button] Clicked - using Univer's native selection API");

                  try {
                    if (!univerAPIRef.current) {
                      alert("Editor not ready. Please try again.");
                      return;
                    }

                    // Get the active document from Univer API
                    const activeDoc = univerAPIRef.current.getActiveDocument();
                    if (!activeDoc) {
                      alert("No active document found.");
                      return;
                    }

                    // Get document snapshot which contains dataStream
                    const snapshot = activeDoc.getSnapshot();
                    if (!snapshot?.body?.dataStream) {
                      alert("Document data not available.");
                      return;
                    }

                    console.log("[AI Button] Document snapshot retrieved");

                    // Get selection from Univer - use univerAPIRef
                    if (!univerInstanceRef.current) {
                      alert("Editor not ready.");
                      return;
                    }

                    const injector = (univerInstanceRef.current as any).__getInjector?.();
                    if (!injector) {
                      alert("Cannot access editor services.");
                      return;
                    }

                    const selectionManager = injector.get(DocSelectionManagerService);
                    if (!selectionManager) {
                      alert("Selection service not available.");
                      return;
                    }

                    const selection = selectionManager.getActiveTextRange();
                    if (!selection) {
                      alert("Please select some text in the document first.");
                      return;
                    }

                    // Extract selected text from dataStream
                    const { startOffset, endOffset } = selection;
                    const dataStream = snapshot.body.dataStream;
                    const selectedText = dataStream.substring(startOffset, endOffset);

                    console.log("[AI Button] Selected text:", selectedText);
                    console.log("[AI Button] Selection range:", { startOffset, endOffset });

                    if (selectedText?.trim()) {
                      // Dispatch AI action event
                      const event = new CustomEvent("univer:ai-action", {
                        detail: {
                          action: "ai-assist",
                          selectedText: selectedText,
                          documentId: activeDoc.getId(),
                          selectionRange: { startOffset, endOffset },
                          timestamp: Date.now(),
                        },
                      });
                      window.dispatchEvent(event);
                      console.log("[AI Button] ✓ AI action event dispatched");
                    } else {
                      alert("Please select some text in the document first.");
                    }
                  } catch (err) {
                    console.error("[AI Button] Error getting selection:", err);
                    alert(
                      "Unable to capture selected text: " +
                        (err instanceof Error ? err.message : String(err)),
                    );
                  }
                }}
                disabled={loading || !!error}
                className="bg-gradient-to-r from-purple-500 to-indigo-600 hover:from-purple-600 hover:to-indigo-700 text-white border-0"
                title="AI Assistant - Select text and click to get AI help"
              >
                <svg
                  className="mr-2 h-4 w-4"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                >
                  <path d="M12 2a2 2 0 0 1 2 2c0 .74-.4 1.39-1 1.73V7h1a7 7 0 0 1 7 7h1a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-1v1a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-1H2a1 1 0 0 1-1-1v-3a1 1 0 0 1 1-1h1a7 7 0 0 1 7-7h1V5.73c-.6-.34-1-.99-1-1.73a2 2 0 0 1 2-2M7.5 13a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3m9 0a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3M8 18h8" />
                </svg>
                AI Assistant
              </Button>
            </div>
          </div>
        </div>
      )}

      {error ? (
        <div className="flex h-full flex-col items-center justify-center p-4">
          <h3 className="mb-2 font-semibold text-lg">
            {error.includes("import")
              ? "Import Failed"
              : "Unable to initialize editor"}
          </h3>
          <p className="mb-4 text-center text-muted-foreground text-sm">
            {error}
          </p>
          {error.includes("import") && (
            <button
              type="button"
              onClick={() => setError(null)}
              className="rounded-md bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700"
            >
              Try Again
            </button>
          )}
        </div>
      ) : (
        <>
          {(loading || importing) && (
            <div className="absolute inset-0 z-10 flex items-center justify-center bg-background/50">
              <div className="text-center">
                <div className="mx-auto mb-2 h-8 w-8 animate-spin rounded-full border-primary border-b-2" />
                <p className="text-muted-foreground text-sm">
                  {importing
                    ? "Importing document..."
                    : "Initializing editor..."}
                </p>
              </div>
            </div>
          )}
          <div
            key={documentData?.id || "blank"}
            className="flex-1 bg-gray-200"
            ref={containerRef}
            style={{
              overflow: "auto",
              position: "relative",
              backgroundColor: "#e5e7eb", /* Gray background to show page boundaries */
            }}
          />
        </>
      )}

      {/* AI Assistant Sheet */}
      <Sheet 
        open={showAISheet} 
        onOpenChange={(open) => {
          console.log("[AI Sheet] State changing to:", open);
          setShowAISheet(open);
        }}
      >
        <SheetContent side="right" className="w-[400px] sm:w-[540px] p-0">
          <div className="p-6">
            <SheetHeader>
              <SheetTitle className="flex items-center gap-2 text-xl font-bold">
                <svg
                  className="h-5 w-5 text-purple-600"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                >
                  <path d="M12 2a2 2 0 0 1 2 2c0 .74-.4 1.39-1 1.73V7h1a7 7 0 0 1 7 7h1a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-1v1a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-1H2a1 1 0 0 1-1-1v-3a1 1 0 0 1 1-1h1a7 7 0 0 1 7-7h1V5.73c-.6-.34-1-.99-1-1.73a2 2 0 0 1 2-2M7.5 13a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3m9 0a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3M8 18h8" />
                </svg>
                AI Assistant
              </SheetTitle>
              <SheetDescription className="text-sm">
                Refine or get help with your selected text
              </SheetDescription>
            </SheetHeader>
            <div className="mt-8 space-y-6">
              <div className="space-y-3">
                <label className="text-sm font-semibold text-foreground">
                  Selected Text
                </label>
                <div className="relative">
                  <textarea
                    value={selectedText}
                    readOnly
                    className="w-full min-h-[140px] p-4 text-sm border border-gray-300 rounded-lg resize-none bg-gray-50 focus:outline-none focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                    placeholder="Select text in the document to get started..."
                  />
                  {selectedText && (
                    <div className="absolute top-2 right-2 text-xs text-gray-500 bg-white px-2 py-0.5 rounded-md shadow-sm">
                      {selectedText.length} chars
                    </div>
                  )}
                </div>
              </div>
              <div className="space-y-3">
                <label className="text-sm font-semibold text-foreground">
                  What would you like to do?
                </label>
                <textarea
                  value={aiPrompt}
                  onChange={(e) => setAiPrompt(e.target.value)}
                  className="w-full min-h-[140px] p-4 text-sm border border-gray-300 rounded-lg resize-none bg-white focus:outline-none focus:ring-2 focus:ring-purple-500 focus:border-transparent"
                  placeholder="e.g., Summarize this text, Improve the writing, Translate to Spanish..."
                />
                <div className="flex flex-wrap gap-2 mt-2">
                  <button
                    type="button"
                    onClick={() => setAiPrompt("Summarize this text")}
                    className="text-xs px-3 py-1.5 rounded-md bg-purple-50 hover:bg-purple-100 text-purple-700 font-medium transition-colors"
                  >
                    Summarize
                  </button>
                  <button
                    type="button"
                    onClick={() => setAiPrompt("Improve the writing")}
                    className="text-xs px-3 py-1.5 rounded-md bg-purple-50 hover:bg-purple-100 text-purple-700 font-medium transition-colors"
                  >
                    Improve
                  </button>
                  <button
                    type="button"
                    onClick={() => setAiPrompt("Translate to Spanish")}
                    className="text-xs px-3 py-1.5 rounded-md bg-purple-50 hover:bg-purple-100 text-purple-700 font-medium transition-colors"
                  >
                    Translate
                  </button>
                  <button
                    type="button"
                    onClick={() => setAiPrompt("Explain this")}
                    className="text-xs px-3 py-1.5 rounded-md bg-purple-50 hover:bg-purple-100 text-purple-700 font-medium transition-colors"
                  >
                    Explain
                  </button>
                </div>
              </div>
              <Button
                className="w-full bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white h-12 text-base font-semibold shadow-md hover:shadow-lg transition-shadow"
                onClick={() => {
                  if (!selectedText.trim()) {
                    alert("Please select some text first");
                    return;
                  }
                  if (!aiPrompt.trim()) {
                    alert("Please enter what you'd like to do");
                    return;
                  }
                  console.log("[AI] Selected:", selectedText);
                  console.log("[AI] Prompt:", aiPrompt);
                  alert(
                    `AI Processing...\n\nSelected: "${selectedText.substring(
                      0,
                      100,
                    )}${
                      selectedText.length > 100 ? "..." : ""
                    }"\n\nPrompt: "${aiPrompt}"\n\n(AI integration would happen here)`,
                  );
                }}
              >
                <svg
                  className="mr-2 h-4 w-4"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2"
                >
                  <path d="M12 2a2 2 0 0 1 2 2c0 .74-.4 1.39-1 1.73V7h1a7 7 0 0 1 7 7h1a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1h-1v1a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-1H2a1 1 0 0 1-1-1v-3a1 1 0 0 1 1-1h1a7 7 0 0 1 7-7h1V5.73c-.6-.34-1-.99-1-1.73a2 2 0 0 1 2-2M7.5 13a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3m9 0a1.5 1.5 0 1 0 0 3 1.5 1.5 0 0 0 0-3M8 18h8" />
                </svg>
                Process with AI
              </Button>
            </div>
          </div>
        </SheetContent>
      </Sheet>
    </div>
  );
}
