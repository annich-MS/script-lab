import {Utilities} from '../helpers';

export enum ContextType {
    Unknown,
    Excel,
    Word,
    PowerPoint,
    OneNote,
    Fabric
}

export class ContextUtil {
    static officeJsUrl = '//appsforoffice.microsoft.com/lib/1/hosted/office.js';
    
    /** Indicates whether the getScript for Office.js has been initiated already */
    static windowkey_initiatedOfficeLoading = 'initiatedOfficeLoading';

    /** Returns true after Office.initialized has been called */
    static windowkey_officeInitialized = 'officeInitialized';

    static sessionStorageKey_context = 'context'
    static sessionStorageKey_wasLaunchedFromAddin = 'wasLaunchedFromAddin';


    static getGlobalState(sessionStorageKey: string) {
        return window[sessionStorageKey];
    }

    static setGlobalState(sessionStorageKey: string, value: any) {
        return window[sessionStorageKey] = value;
    }

    static get contextString(): string {
        return window.sessionStorage.getItem(ContextUtil.sessionStorageKey_context);
    }

    static get isAddin(): boolean {
        // Note: it's an intentional string comparison.
        return window.sessionStorage.getItem(ContextUtil.sessionStorageKey_wasLaunchedFromAddin) === 'true';
    }

    /** 
     * Gets the context type or "unknown".  Note, this function does NOT throw on unknown,
     * though many of the derived ones (getContextNamespace, contextTagline, etc.) do.
     */
    static get context(): ContextType {
        switch (ContextUtil.contextString) {
            case 'excel':
                return ContextType.Excel;
            case 'word':
                return ContextType.Word;
            case 'powerpoint':
                return ContextType.PowerPoint;
            case 'onenote':
                return ContextType.Word;
            case 'fabric':
                return ContextType.Fabric;
            default:
                return ContextType.Unknown;
        }
    }

    static get hostName() {
        switch (ContextUtil.context) {
            case ContextType.Excel:
                return 'Excel';
            case ContextType.Word:
                return 'Word';
            case ContextType.PowerPoint:
                return 'PowerPoint'
            case ContextType.OneNote:
                return 'OneNote';
            default:
                throw new Error("Invalid context type for Office namespace");
        }
    }

    static get contextNamespace() {
        switch (ContextUtil.context) {
            case ContextType.Excel:
                return 'Excel';
            case ContextType.Word:
                return 'Word';
            case ContextType.PowerPoint:
                return null; // Intentionally missing until PowerPoint has the new host-specific API model
            case ContextType.OneNote:
                return 'OneNote';
            default:
                throw new Error("Invalid context type for Office namespace");
        }
    }

    static get contextTagline(): string {
        switch (ContextUtil.context) {
            case ContextType.Excel:
            case ContextType.Word:
            case ContextType.PowerPoint:
            case ContextType.OneNote:
                return 'Office Add-in Playground';

            case ContextType.Fabric:
                return 'Fabric Playground';

            default: 
                throw new Error("Cannot determine playground context");
        }
    }

    static get fullPlaygroundDescription(): string {
        switch (ContextUtil.context) {
            case ContextType.Excel:
            case ContextType.Word:
            case ContextType.PowerPoint:
            case ContextType.OneNote:
                return "Office Add-in Playground - " + ContextUtil.hostName;

            case ContextType.Fabric:
                return "Fabric Playground";

            default:
                throw "Invalid context " + ContextUtil.context;
        }
    }

    static get isOfficeContext(): boolean {
        switch (ContextUtil.context) {
            case ContextType.Excel:
            case ContextType.Word:
            case ContextType.PowerPoint:
            case ContextType.OneNote:
                return true;

            default:
                return false;
        }
    }
}