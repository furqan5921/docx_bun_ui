/**
 * Custom Quick Insert Plugin for Univer Docs
 *
 * This plugin provides a slash command menu with AI button functionality.
 * Based on @univerjs/docs-quick-insert-ui but with custom AI integration.
 */

import type { DocumentDataModel, IDisposable, Nullable } from "@univerjs/core";
import {
  CommandType,
  Disposable,
  ICommandService,
  Inject,
  Injector,
  IUniverInstanceService,
  Plugin,
  RANGE_DIRECTION,
  UniverInstanceType,
} from "@univerjs/core";
import { DocSelectionManagerService } from "@univerjs/docs";
import {
  CutContentCommand,
  DocCanvasPopManagerService,
  DocEventManagerService,
  type IInnerCutCommandParams,
} from "@univerjs/docs-ui";
import { IRenderManagerService } from "@univerjs/engine-render";
import type { Observable } from "rxjs";
import { BehaviorSubject, distinctUntilChanged, map } from "rxjs";

// ============================================================================
// Types and Interfaces
// ============================================================================

export interface IQuickInsertMenuItem {
  id: string;
  icon?: string;
  title: string;
  keywords?: string[];
  isAI?: boolean; // Flag for AI menu item
}

export interface IQuickInsertMenuGroup {
  id: string;
  icon?: string;
  title: string;
  children?: IQuickInsertMenuItem[];
}

export type QuickInsertMenu = IQuickInsertMenuGroup | IQuickInsertMenuItem;

// ============================================================================
// Commands
// ============================================================================

interface IDeleteSearchKeyCommandParams {
  start: number;
  end: number;
}

export const DeleteSearchKeyCommand = {
  id: "custom.doc.command.delete-search-key",
  type: CommandType.COMMAND,
  handler: (accessor: any, params: IDeleteSearchKeyCommandParams) => {
    const commandService = accessor.get(ICommandService);
    const { start, end } = params;
    return commandService.syncExecuteCommand(CutContentCommand.id, {
      segmentId: "",
      textRanges: [
        {
          startOffset: start,
          endOffset: start,
          collapsed: true,
        },
      ],
      selections: [
        {
          startOffset: start,
          endOffset: end,
          collapsed: false,
          direction: RANGE_DIRECTION.FORWARD,
        },
      ],
    } as IInnerCutCommandParams);
  },
};

const noopDisposable = {
  dispose: () => {},
};

// ============================================================================
// Custom Quick Insert Service
// ============================================================================

export class CustomQuickInsertService extends Disposable {
  private readonly _editPopup$ = new BehaviorSubject<
    Nullable<{
      anchor: number;
      disposable: IDisposable;
      unitId: string;
    }>
  >(undefined);

  readonly editPopup$ = this._editPopup$.asObservable();
  get editPopup() {
    return this._editPopup$.value;
  }

  private readonly _isComposing$ = new BehaviorSubject<boolean>(false);
  readonly isComposing$ = this._isComposing$.asObservable();
  get isComposing() {
    return this._isComposing$.value;
  }

  setIsComposing(isComposing: boolean) {
    this._isComposing$.next(isComposing);
  }

  private readonly _inputOffset$ = new BehaviorSubject<{
    start: number;
    end: number;
  }>({
    start: 0,
    end: 0,
  });

  readonly inputOffset$ = this._inputOffset$.asObservable();
  get inputOffset() {
    return this._inputOffset$.value;
  }

  setInputOffset(offset: { start: number; end: number }) {
    this._inputOffset$.next(offset);
  }

  readonly filterKeyword$: Observable<string>;

  private _menuSelectedCallbacks: Set<(menu: IQuickInsertMenuItem) => void> =
    new Set();

  private _currentMenus: QuickInsertMenu[] = [];

  private _docCanvasPopupManagerService: DocCanvasPopManagerService;
  private _univerInstanceService: IUniverInstanceService;
  private _commandService: ICommandService;
  private _renderManagerService: IRenderManagerService;
  private _docSelectionManagerService: DocSelectionManagerService;

  constructor(
    docCanvasPopupManagerService: DocCanvasPopManagerService,
    univerInstanceService: IUniverInstanceService,
    commandService: ICommandService,
    renderManagerService: IRenderManagerService,
    docSelectionManagerService: DocSelectionManagerService,
  ) {
    super();

    this._docCanvasPopupManagerService = docCanvasPopupManagerService;
    this._univerInstanceService = univerInstanceService;
    this._commandService = commandService;
    this._renderManagerService = renderManagerService;
    this._docSelectionManagerService = docSelectionManagerService;

    this.disposeWithMe(this._editPopup$);

    const getBodySlice = (start: number, end: number) =>
      this._univerInstanceService
        .getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC)
        ?.getBody()
        ?.dataStream.slice(start, end);

    this.filterKeyword$ = this._inputOffset$.pipe(
      map((offset) => {
        const slice = getBodySlice(offset.start, offset.end);
        return slice?.slice(1) ?? "";
      }),
      distinctUntilChanged(),
    );

    // Register the delete command
    this._commandService.registerCommand(DeleteSearchKeyCommand);
  }

  setMenus(menus: QuickInsertMenu[]) {
    this._currentMenus = menus;
  }

  getMenus(): QuickInsertMenu[] {
    return this._currentMenus;
  }

  private getDocEventManagerService(unitId: string) {
    return this._renderManagerService
      .getRenderById(unitId)
      ?.with(DocEventManagerService);
  }

  showPopup(options: { index: number; unitId: string }) {
    console.log("[CustomQuickInsertService] showPopup called with:", options);
    const { index, unitId } = options;
    this.closePopup();

    const currentDoc =
      this._univerInstanceService.getUnit<DocumentDataModel>(unitId);
    const paragraphs = currentDoc?.getBody()?.paragraphs;

    console.log(
      "[CustomQuickInsertService] Paragraphs:",
      paragraphs?.length,
      "First paragraph:",
      paragraphs?.[0],
    );

    if (!paragraphs || paragraphs.length === 0) {
      console.warn("[CustomQuickInsertService] No paragraphs in document");
      return;
    }

    // Find the paragraph that CONTAINS or is CLOSEST TO the index
    let paragraph: any = null;
    
    // Special case: if index is before first paragraph, use first paragraph
    if (paragraphs.length > 0 && index < paragraphs[0].startIndex) {
      paragraph = paragraphs[0];
      console.log(
        `[CustomQuickInsertService] ✓ Index ${index} is before first paragraph (${paragraphs[0].startIndex}), using first paragraph`,
      );
    } else {
      // Find the paragraph that contains the index
      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const nextP = paragraphs[i + 1];
        const pStart = p.startIndex;
        // For the last paragraph, accept any index >= pStart (including end of file)
        const pEnd = nextP ? nextP.startIndex : Number.MAX_SAFE_INTEGER;

        console.log(
          `[CustomQuickInsertService] Checking paragraph ${i}: range [${pStart}-${pEnd}), index=${index}`,
        );

        // Use <= for the last paragraph to include positions at or after paragraph start
        const isInRange = nextP ? (index >= pStart && index < pEnd) : (index >= pStart);
        
        if (isInRange) {
          paragraph = p;
          console.log(
            `[CustomQuickInsertService] ✓ Found matching paragraph ${i}`,
          );
          break;
        }
      }
    }

    if (!paragraph) {
      console.warn(
        "[CustomQuickInsertService] No paragraph found containing index",
        index,
      );
      return;
    }

    console.log(
      "[CustomQuickInsertService] Found paragraph:",
      paragraph.startIndex,
      "containing index",
      index,
    );

    const docEventManagerService = this.getDocEventManagerService(unitId);
    console.log(
      "[CustomQuickInsertService] docEventManagerService exists:",
      !!docEventManagerService,
    );

    if (!docEventManagerService) {
      console.warn(
        "[CustomQuickInsertService] DocEventManagerService not available",
      );
      return;
    }

    let bounds: any = null;

    // Use paragraph bounds and estimate offset based on character distance
    {
      let paragraphBound =
        docEventManagerService.findParagraphBoundByIndex(index);

      console.log(
        "[CustomQuickInsertService] paragraphBound from cursor index:",
        paragraphBound,
      );

      if (!paragraphBound) {
        // Try with paragraph start index
        paragraphBound = docEventManagerService.findParagraphBoundByIndex(
          paragraph.startIndex,
        );

        console.log(
          "[CustomQuickInsertService] paragraphBound from paragraph start:",
          paragraphBound,
        );
      }

      if (paragraphBound) {
        bounds = { ...paragraphBound.firstLine };
        
        // Estimate horizontal offset based on character distance from paragraph start
        // Assume average character width of ~8 pixels (will vary by font/size)
        const charOffset = index - paragraph.startIndex;
        const estimatedOffset = charOffset * 8;
        
        bounds.left = bounds.left + estimatedOffset;
        bounds.right = bounds.left + 10;
        
        console.log(
          "[CustomQuickInsertService] ✓ Using estimated position with offset:",
          charOffset,
          "chars =",
          estimatedOffset,
          "px, bounds:",
          bounds,
        );
      }
    }

    if (!bounds) {
      console.warn(
        "[CustomQuickInsertService] Could not determine position for popup menu",
      );
      return;
    }

    console.log("[CustomQuickInsertService] Final bounds for popup:", bounds);

    // Emit event for React to show the popup
    const event = new CustomEvent("univer:quick-insert-show", {
      detail: {
        anchor: index,
        unitId,
        bounds: bounds,
      },
    });
    window.dispatchEvent(event);

    // Store popup state
    const disposable = {
      dispose: () => {
        const closeEvent = new CustomEvent("univer:quick-insert-close");
        window.dispatchEvent(closeEvent);
      },
    };

    this._editPopup$.next({ disposable, anchor: index, unitId });
    console.log("[CustomQuickInsertService] Popup state updated");
  }

  closePopup() {
    if (this.editPopup) {
      this.editPopup.disposable.dispose();
      this._editPopup$.next(null);
    }
  }

  onMenuSelected(callback: (menu: IQuickInsertMenuItem) => void) {
    this._menuSelectedCallbacks.add(callback);
    return () => {
      this._menuSelectedCallbacks.delete(callback);
    };
  }

  emitMenuSelected(menu: IQuickInsertMenuItem) {
    console.log(
      "[CustomQuickInsertService] ========================================",
    );
    console.log("[CustomQuickInsertService] emitMenuSelected called");
    console.log("[CustomQuickInsertService] Menu:", menu);
    console.log("[CustomQuickInsertService] Input offset:", this.inputOffset);
    console.log(
      "[CustomQuickInsertService] Callbacks count:",
      this._menuSelectedCallbacks.size,
    );

    const { start, end } = this.inputOffset;

    // Delete the search key first
    console.log(
      `[CustomQuickInsertService] Deleting search key from ${start} to ${end}`,
    );
    try {
      const result = this._commandService.syncExecuteCommand(
        DeleteSearchKeyCommand.id,
        {
          start,
          end,
        },
      );
      console.log("[CustomQuickInsertService] Delete command result:", result);
    } catch (error) {
      console.error(
        "[CustomQuickInsertService] Error deleting search key:",
        error,
      );
    }

    // Then emit the menu selection
    setTimeout(() => {
      console.log(
        "[CustomQuickInsertService] Calling menu selection callbacks",
      );
      this._menuSelectedCallbacks.forEach((callback) => {
        console.log("[CustomQuickInsertService] Calling callback...");
        callback(menu);
      });

      // Emit custom event for React to handle
      const event = new CustomEvent("univer:quick-insert-selected", {
        detail: { menu },
      });
      window.dispatchEvent(event);
      console.log(
        "[CustomQuickInsertService] Dispatched univer:quick-insert-selected event",
      );
      console.log(
        "[CustomQuickInsertService] ========================================",
      );
    }, 0);
  }
}

// ============================================================================
// Custom Quick Insert Plugin
// ============================================================================

export class CustomQuickInsertPlugin extends Plugin {
  static override pluginName = "CUSTOM_QUICK_INSERT_PLUGIN";
  static override type = UniverInstanceType.UNIVER_DOC;

  private _service: CustomQuickInsertService | null = null;

  constructor(
    _config: unknown,
    @Inject(Injector) protected readonly _injector: Injector,
  ) {
    super();
  }

  override onStarting(): void {
    // Register dependencies - this happens before onReady
    // We'll create the service in onReady when injector is available
  }

  override onReady(): void {
    // Get injector from the instance
    const injector = this._injector;

    // Create and store the service manually
    this._service = new CustomQuickInsertService(
      injector.get(DocCanvasPopManagerService),
      injector.get(IUniverInstanceService),
      injector.get(ICommandService),
      injector.get(IRenderManagerService),
      injector.get(DocSelectionManagerService),
    );

    // Store on window for React to access
    if (typeof window !== "undefined") {
      (window as any).__customQuickInsertService = this._service;
    }

    // Register the delete command
    const commandService = injector.get(ICommandService);
    try {
      commandService.registerCommand(DeleteSearchKeyCommand);
    } catch (e: any) {
      // Command might already be registered (e.g., in HMR or hot reload)
      // This is safe to ignore - the command is already available
      if (e?.message?.includes("has been registered before")) {
        console.log(
          "[CustomQuickInsertPlugin] Command already registered (HMR), skipping",
        );
      } else {
        console.warn(
          "[CustomQuickInsertPlugin] Command registration error:",
          e?.message,
        );
      }
    }

    this._setupPlugin();
  }

  private _setupPlugin(): void {
    if (!this._service) {
      console.error("[CustomQuickInsertPlugin] Service not initialized");
      return;
    }

    const service = this._service;

    // Define default menus with AI button
    const defaultMenus: QuickInsertMenu[] = [
      {
        id: "ai-group",
        title: "AI",
        icon: "✨",
        children: [
          {
            id: "ask-ai",
            title: "Ask AI",
            icon: "🤖",
            keywords: ["ai", "assistant", "help", "ask"],
            isAI: true,
          },
        ],
      },
      {
        id: "basic-group",
        title: "Basic",
        children: [
          {
            id: "heading-1",
            title: "Heading 1",
            icon: "H1",
            keywords: ["heading", "h1", "title"],
          },
          {
            id: "heading-2",
            title: "Heading 2",
            icon: "H2",
            keywords: ["heading", "h2", "subtitle"],
          },
          {
            id: "heading-3",
            title: "Heading 3",
            icon: "H3",
            keywords: ["heading", "h3"],
          },
          {
            id: "paragraph",
            title: "Paragraph",
            icon: "¶",
            keywords: ["paragraph", "text", "p"],
          },
        ],
      },
      {
        id: "blocks-group",
        title: "Blocks",
        children: [
          {
            id: "table",
            title: "Table",
            icon: "⊞",
            keywords: ["table", "grid"],
          },
        ],
      },
    ];

    service.setMenus(defaultMenus);

    // Setup menu selection handler
    service.onMenuSelected((menu) => {
      console.log(
        "[CustomQuickInsertPlugin] ========================================",
      );
      console.log("[CustomQuickInsertPlugin] Menu selected:", menu);
      console.log("[CustomQuickInsertPlugin] Menu ID:", menu.id);
      console.log("[CustomQuickInsertPlugin] Menu title:", menu.title);
      console.log("[CustomQuickInsertPlugin] Is AI?", menu.isAI);
      console.log(
        "[CustomQuickInsertPlugin] ========================================",
      );

      if (menu.isAI) {
        console.log("[CustomQuickInsertPlugin] Handling as AI selection");
        // Handle AI menu item
        this._handleAISelection();
      } else {
        console.log("[CustomQuickInsertPlugin] Handling as block insertion");
        // Handle other menu items
        this._handleBlockInsertion(menu);
      }
    });

    // Listen for "/" key to show popup
    this._setupSlashCommandListener();
  }

  private _handleAISelection(): void {
    const selectionManager = this._injector.get(DocSelectionManagerService);
    const selection = selectionManager.getActiveTextRange();

    if (!selection) {
      console.warn("[CustomQuickInsertPlugin] No active selection");
      return;
    }

    const univerInstanceService = this._injector.get(IUniverInstanceService);
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );

    if (!currentDoc) {
      console.warn("[CustomQuickInsertPlugin] No active document");
      return;
    }

    const body = currentDoc.getBody();
    const selectedText =
      body?.dataStream.substring(selection.startOffset, selection.endOffset) ||
      "";

    // Emit AI action event
    const event = new CustomEvent("univer:ai-action", {
      detail: {
        action: "ai-assist",
        selectedText,
        documentId: currentDoc.getUnitId(),
        selectionRange: {
          startOffset: selection.startOffset,
          endOffset: selection.endOffset,
          collapsed: selection.collapsed,
        },
      },
    });
    window.dispatchEvent(event);
  }

  private async _handleBlockInsertion(
    menu: IQuickInsertMenuItem,
  ): Promise<void> {
    console.log(
      "[CustomQuickInsertPlugin] _handleBlockInsertion called with menu:",
      menu,
    );

    const univerInstanceService = this._injector.get(IUniverInstanceService);
    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );

    console.log("[CustomQuickInsertPlugin] Current doc exists:", !!currentDoc);
    if (!currentDoc) {
      console.error("[CustomQuickInsertPlugin] No current document found!");
      return;
    }

    const commandService = this._injector.get(ICommandService);
    const selectionManager = this._injector.get(DocSelectionManagerService);

    console.log(
      "[CustomQuickInsertPlugin] Got commandService:",
      !!commandService,
    );
    console.log(
      "[CustomQuickInsertPlugin] Got selectionManager:",
      !!selectionManager,
    );

    // Get current selection/cursor position
    const selection = selectionManager.getActiveTextRange();
    console.log("[CustomQuickInsertPlugin] Current selection:", selection);

    if (!selection) {
      console.error(
        "[CustomQuickInsertPlugin] No active selection/cursor position!",
      );
      return;
    }

    // Handle different block types
    console.log("[CustomQuickInsertPlugin] Processing block type:", menu.id);

    try {
      switch (menu.id) {
        case "heading-1":
          console.log("[CustomQuickInsertPlugin] Inserting Heading 1");
          await this._insertHeading(1);
          break;
        case "heading-2":
          console.log("[CustomQuickInsertPlugin] Inserting Heading 2");
          await this._insertHeading(2);
          break;
        case "heading-3":
          console.log("[CustomQuickInsertPlugin] Inserting Heading 3");
          await this._insertHeading(3);
          break;
        case "paragraph":
          console.log("[CustomQuickInsertPlugin] Inserting Paragraph");
          await this._insertParagraph();
          break;
        case "table":
          console.log("[CustomQuickInsertPlugin] Inserting Table");
          await this._insertTable();
          break;
        default:
          console.warn(
            "[CustomQuickInsertPlugin] Unknown block type:",
            menu.id,
          );
      }
      console.log("[CustomQuickInsertPlugin] ✅ Block insertion completed");
    } catch (error) {
      console.error(
        "[CustomQuickInsertPlugin] ❌ Error during block insertion:",
        error,
      );
    }

    // Emit block insertion event for React components
    const event = new CustomEvent("univer:insert-block", {
      detail: {
        blockType: menu.id,
        documentId: currentDoc.getUnitId(),
      },
    });
    window.dispatchEvent(event);
    console.log("[CustomQuickInsertPlugin] Emitted univer:insert-block event");
  }

  private async _insertHeading(level: 1 | 2 | 3): Promise<void> {
    console.log(`[CustomQuickInsertPlugin] _insertHeading(${level}) called`);

    const commandService = this._injector.get(ICommandService);
    const selectionManager = this._injector.get(DocSelectionManagerService);
    const univerInstanceService = this._injector.get(IUniverInstanceService);

    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );

    if (!currentDoc) {
      console.error(
        "[CustomQuickInsertPlugin] No current document in _insertHeading",
      );
      return;
    }

    const selection = selectionManager.getActiveTextRange();
    if (!selection) {
      console.error("[CustomQuickInsertPlugin] No selection in _insertHeading");
      return;
    }

    console.log(
      `[CustomQuickInsertPlugin] Current cursor position:`,
      selection.startOffset,
    );

    // Insert placeholder text for the heading
    const headingText = `Heading ${level}\n`;
    console.log(`[CustomQuickInsertPlugin] Inserting text: "${headingText}"`);

    try {
      const result = await commandService.executeCommand(
        "doc.command.insert-text",
        {
          unitId: currentDoc.getUnitId(),
          body: {
            dataStream: headingText,
          },
          range: {
            startOffset: selection.startOffset,
            endOffset: selection.startOffset,
            collapsed: true,
          },
          segmentId: "",
        },
      );

      console.log(
        `[CustomQuickInsertPlugin] Command execution result:`,
        result,
      );

      if (!result) {
        console.error(
          `[CustomQuickInsertPlugin] ❌ Command returned false - insertion failed`,
        );
      } else {
        console.log(
          `[CustomQuickInsertPlugin] ✅ Heading ${level} inserted successfully`,
        );
      }
    } catch (error) {
      console.error(
        `[CustomQuickInsertPlugin] ❌ Error inserting heading:`,
        error,
      );
    }
  }

  private async _insertTable(): Promise<void> {
    console.log("[CustomQuickInsertPlugin] _insertTable() called");
    
    // Table insertion is not supported in Univer preset-docs-core v0.15.x
    // The table functionality exists for importing from DOCX, but programmatic 
    // insertion through slash commands requires @univerjs/docs-table plugin
    // which is not available in this version.
    
    // Instead, insert a placeholder text that indicates where a table should be
    const commandService = this._injector.get(ICommandService);
    const selectionManager = this._injector.get(DocSelectionManagerService);
    const univerInstanceService = this._injector.get(IUniverInstanceService);

    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );

    if (!currentDoc) {
      console.error(
        "[CustomQuickInsertPlugin] No current document in _insertTable",
      );
      return;
    }

    const selection = selectionManager.getActiveTextRange();
    if (!selection) {
      console.error("[CustomQuickInsertPlugin] No selection in _insertTable");
      return;
    }

    console.log(
      "[CustomQuickInsertPlugin] ⚠️ Table insertion via slash command not supported in this Univer version",
    );
    console.log(
      "[CustomQuickInsertPlugin] Inserting placeholder text instead. Tables can be imported from DOCX files.",
    );
    
    // Insert a placeholder that user can replace
    const placeholderText = "[Table - Import from DOCX or use external editor]\n";
    
    try {
      const result = await commandService.executeCommand(
        "doc.command.insert-text",
        {
          unitId: currentDoc.getUnitId(),
          body: {
            dataStream: placeholderText,
          },
          range: {
            startOffset: selection.startOffset,
            endOffset: selection.startOffset,
            collapsed: true,
          },
          segmentId: "",
        },
      );

      if (result) {
        console.log("[CustomQuickInsertPlugin] ✅ Placeholder inserted successfully");
      } else {
        console.error("[CustomQuickInsertPlugin] ❌ Failed to insert placeholder");
      }
    } catch (error) {
      console.error("[CustomQuickInsertPlugin] ❌ Error inserting placeholder:", error);
    }
  }

  private async _insertParagraph(): Promise<void> {
    console.log("[CustomQuickInsertPlugin] _insertParagraph() called");

    const commandService = this._injector.get(ICommandService);
    const selectionManager = this._injector.get(DocSelectionManagerService);
    const univerInstanceService = this._injector.get(IUniverInstanceService);

    const currentDoc =
      univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );

    if (!currentDoc) {
      console.error(
        "[CustomQuickInsertPlugin] No current document in _insertParagraph",
      );
      return;
    }

    const selection = selectionManager.getActiveTextRange();
    if (!selection) {
      console.error(
        "[CustomQuickInsertPlugin] No selection in _insertParagraph",
      );
      return;
    }

    console.log(
      "[CustomQuickInsertPlugin] Current cursor position:",
      selection.startOffset,
    );

    // Insert placeholder text for paragraph
    const paragraphText = "Type your paragraph here...\n";
    console.log(`[CustomQuickInsertPlugin] Inserting text: "${paragraphText}"`);

    try {
      const result = await commandService.executeCommand(
        "doc.command.insert-text",
        {
          unitId: currentDoc.getUnitId(),
          body: {
            dataStream: paragraphText,
          },
          range: {
            startOffset: selection.startOffset,
            endOffset: selection.startOffset,
            collapsed: true,
          },
          segmentId: "",
        },
      );

      console.log(
        "[CustomQuickInsertPlugin] Command execution result:",
        result,
      );

      if (!result) {
        console.error(
          "[CustomQuickInsertPlugin] ❌ Command returned false - insertion failed",
        );
      } else {
        console.log(
          "[CustomQuickInsertPlugin] ✅ Paragraph inserted successfully",
        );
      }
    } catch (error) {
      console.error(
        "[CustomQuickInsertPlugin] ❌ Error inserting paragraph:",
        error,
      );
    }
  }

  private _setupSlashCommandListener(): void {
    if (!this._service) {
      console.error(
        "[CustomQuickInsertPlugin] Service not available for slash command listener",
      );
      return;
    }

    const service = this._service;
    const selectionManager = this._injector.get(DocSelectionManagerService);
    const univerInstanceService = this._injector.get(IUniverInstanceService);
    const commandService = this._injector.get(ICommandService);

    console.log("[CustomQuickInsertPlugin] Setting up slash command listener");

    let isMenuOpen = false;
    let lastSlashPos = -1;
    let debounceTimeout: number | null = null;

    // NOTE: We DO NOT intercept keyboard events for "/" because it interferes with
    // natural text input and causes "/" to be inserted in the wrong place.
    // Instead, we rely on detecting "/" via text editing commands after Univer
    // has properly inserted it at the cursor position.

    // Subscribe to ALL command executions to detect text changes
    const disposable = commandService.onCommandExecuted((command) => {
      // Only check after text editing commands
      if (
        !command.id.includes("doc.mutation.rich-text-editing") &&
        !command.id.includes("doc.command.insert") &&
        !command.id.includes("doc.command.input") &&
        !command.id.includes("InsertCommand")
      ) {
        return;
      }

      console.log(
        "[CustomQuickInsertPlugin] Text editing command detected:",
        command.id,
      );

      // Clear previous debounce timeout
      if (debounceTimeout !== null) {
        clearTimeout(debounceTimeout);
      }

      // Debounce the detection to avoid rapid checks
      debounceTimeout = window.setTimeout(() => {
        try {
          const currentDoc =
            univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
              UniverInstanceType.UNIVER_DOC,
            );

          if (!currentDoc) return;

          const selection = selectionManager.getActiveTextRange();
          if (!selection || selection.collapsed === false) return;

          const body = currentDoc.getBody();
          if (!body) return;

          const dataStream = body.dataStream;
          const cursorPos = selection.startOffset;

          // Check for "/" pattern - simplified to work everywhere
          let shouldShowMenu = false;
          let slashPos = -1;

          if (cursorPos > 0) {
            // Look back in the text before cursor to find the most recent "/"
            // But only within a reasonable distance (20 chars) to avoid false triggers
            const searchStart = Math.max(0, cursorPos - 20);
            const textBefore = dataStream.substring(searchStart, cursorPos);
            const lastSlashIndex = textBefore.lastIndexOf("/");

            if (lastSlashIndex !== -1) {
              slashPos = searchStart + lastSlashIndex;
              const afterSlash = dataStream.substring(slashPos + 1, cursorPos);

              // Allow "/" anywhere - just check that text after slash is on same line
              // and consists of valid filter characters (letters, numbers, or empty)
              const isValidFilter = /^[a-zA-Z0-9]*$/.test(afterSlash);
              const isOnSameLine = !afterSlash.includes("\r") && !afterSlash.includes("\n");
              const isReasonablyClose = afterSlash.length <= 20; // Close to slash

              // Additional check: slash must be recent (within last 20 chars of cursor)
              const slashIsRecent = (cursorPos - slashPos) <= 20;

              if (isValidFilter && isOnSameLine && isReasonablyClose && slashIsRecent) {
                shouldShowMenu = true;
                console.log(
                  "[CustomQuickInsertPlugin] ✅ Valid slash detected at position",
                  slashPos,
                  "with filter:",
                  JSON.stringify(afterSlash),
                );

                // Update filter keyword offset
                service.setInputOffset({
                  start: slashPos,
                  end: cursorPos,
                });
              }
            }
          }

          // Show or update menu
          if (shouldShowMenu && slashPos !== -1) {
            if (!isMenuOpen || lastSlashPos !== slashPos) {
              console.log(
                `[CustomQuickInsertPlugin] "/" detected at position ${slashPos}, cursor at ${cursorPos}`,
              );
              isMenuOpen = true;
              lastSlashPos = slashPos;
              service.showPopup({
                index: slashPos,
                unitId: currentDoc.getUnitId(),
              });
            } else {
              // Menu already open, just update the filter
              service.setInputOffset({
                start: slashPos,
                end: cursorPos,
              });
            }
          } else if (isMenuOpen) {
            // Close menu if "/" is gone or cursor moved away (but not if just opened)
            const editPopup = service.editPopup;
            if (editPopup) {
              const startPos = editPopup.anchor;
              if (
                cursorPos < startPos ||
                dataStream[startPos] !== "/" ||
                cursorPos > startPos + 30
              ) {
                console.log(
                  "[CustomQuickInsertPlugin] Closing menu - condition no longer met",
                );
                service.closePopup();
                isMenuOpen = false;
                lastSlashPos = -1;
              }
            }
          }
        } catch (error) {
          console.error(
            "[CustomQuickInsertPlugin] Error in slash detection:",
            error,
          );
        }
      }, 50); // 50ms debounce delay for faster response
    });

    // Cleanup function
    const cleanup = () => {
      disposable.dispose();
    };

    this.disposeWithMe({
      dispose: cleanup,
    });
    
    console.log("[CustomQuickInsertPlugin] Slash command listener active");
  }
}
