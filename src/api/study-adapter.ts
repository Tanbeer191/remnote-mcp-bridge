/**
 * Study-workflow actions: PDF attachment, sources, folders, delete and card reading.
 *
 * Kept separate from RemAdapter so the upstream adapter stays easy to rebase. Uses only the
 * public plugin SDK. Card due dates are not set here: the SDK cannot pass a date and RemNote
 * 1.28 forces all plugins into sandboxed mode, so due dates are set by the companion
 * Chrome-automation script instead (see FINDINGS.md in the parent folder).
 */

import { BuiltInPowerupCodes, type ReactRNPlugin } from '@remnote/plugin-sdk';
import { type AutomationBridgeSettings, SETTING_ACCEPT_REPLACE_OPERATION } from '../settings';

type PluginRemLike = NonNullable<Awaited<ReturnType<ReactRNPlugin['rem']['findOne']>>>;

const UPLOADED_FILES_FOLDER_TITLE = 'Uploaded Files';

export interface AttachPdfParams {
  documentRemId: string;
  url: string;
  fileName: string;
  uploadsFolderRemId?: string;
  dryRun?: boolean;
}

export interface SourceParams {
  remId: string;
  sourceRemId: string;
}

export interface SetFolderStatusParams {
  remId: string;
  isFolder: boolean;
  dryRun?: boolean;
}

export interface DeleteNoteParams {
  remId: string;
  dryRun?: boolean;
  expectedTitle?: string;
}

export interface GetCardsParams {
  remId: string;
  includeHistory?: boolean;
}

export class StudyAdapter {
  constructor(
    private readonly plugin: ReactRNPlugin,
    private readonly getSettings: () => AutomationBridgeSettings
  ) {}

  // ------------------------------------------------------------------ PDFs and sources

  async attachPdf(params: AttachPdfParams) {
    this.requireWrite();
    const documentRemId = requireString(params.documentRemId, 'documentRemId');
    const url = requireString(params.url, 'url');
    const fileName = requireString(params.fileName, 'fileName');
    if (!/^https:\/\//.test(url)) {
      throw new Error('url must be an https URL (e.g. the URL returned by the PDF upload)');
    }
    const dryRun = params.dryRun !== false;

    const doc = await this.findRem(documentRemId);
    const folder = await this.findUploadsFolder(params.uploadsFolderRemId);
    const parentRemId = folder?._id ?? documentRemId;
    const preview = {
      documentRemId,
      documentTitle: await this.title(doc),
      pdfParentRemId: parentRemId,
      pdfParentTitle: folder ? await this.title(folder) : await this.title(doc),
      fileName,
      url,
    };
    if (dryRun) {
      return {
        ...preview,
        dryRun: true,
        created: false,
        pdfRemId: undefined as string | undefined,
      };
    }

    const pdfRem = await this.plugin.rem.createRem();
    if (!pdfRem) {
      throw new Error('RemNote did not create the PDF Rem');
    }
    await pdfRem.setText([fileName]);
    await pdfRem.setParent(parentRemId);
    await pdfRem.addPowerup(BuiltInPowerupCodes.UploadedFile);
    await pdfRem.setPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'URL', [url]);
    await pdfRem.setPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'Name', [fileName]);
    await pdfRem.setIsDocument(true);
    await doc.addSource(pdfRem);

    return { ...preview, dryRun: false, created: true, pdfRemId: pdfRem._id };
  }

  async addSource(params: SourceParams) {
    this.requireWrite();
    const rem = await this.findRem(requireString(params.remId, 'remId'));
    const source = await this.findRem(requireString(params.sourceRemId, 'sourceRemId'));
    await rem.addSource(source);
    return { remId: rem._id, sourceRemIds: await this.sourceIds(rem) };
  }

  async removeSource(params: SourceParams) {
    this.requireWrite();
    const rem = await this.findRem(requireString(params.remId, 'remId'));
    await rem.removeSource(requireString(params.sourceRemId, 'sourceRemId'));
    return { remId: rem._id, sourceRemIds: await this.sourceIds(rem) };
  }

  async getSources(params: { remId: string }) {
    const rem = await this.findRem(requireString(params.remId, 'remId'));
    const sources = await rem.getSources();
    return {
      remId: rem._id,
      sources: await Promise.all(
        sources.map(async (s) => ({
          remId: s._id,
          title: await this.title(s),
          url: await this.uploadedFileUrl(s),
        }))
      ),
    };
  }

  // ------------------------------------------------------------------ structure

  async setFolderStatus(params: SetFolderStatusParams) {
    this.requireWrite();
    const rem = await this.findRem(requireString(params.remId, 'remId'));
    const isFolder = requireBoolean(params.isFolder, 'isFolder');
    const dryRun = params.dryRun !== false;
    const oldIsFolder = await rem.isFolder();
    const result = {
      remId: rem._id,
      title: await this.title(rem),
      oldIsFolder,
      requestedIsFolder: isFolder,
      wouldChange: oldIsFolder !== isFolder,
      dryRun,
    };
    if (dryRun || !result.wouldChange) {
      return { ...result, changed: false };
    }
    await rem.setIsFolder(isFolder);
    return { ...result, changed: true, newIsFolder: await rem.isFolder() };
  }

  async deleteNote(params: DeleteNoteParams) {
    this.requireWrite();
    // Read live from RemNote: the runtime's cached settings only refresh while the sidebar
    // widget is mounted, so a just-ticked box would otherwise be ignored.
    const replaceAllowed =
      (await this.plugin.settings.getSetting<boolean>(SETTING_ACCEPT_REPLACE_OPERATION)) === true;
    if (!replaceAllowed) {
      throw new Error('Delete requires "Accept replace operation" in Automation Bridge settings');
    }
    const rem = await this.findRem(requireString(params.remId, 'remId'));
    const title = await this.title(rem);
    const dryRun = params.dryRun !== false;
    const descendants = await rem.getDescendants();
    const preview = { remId: rem._id, title, descendantCount: descendants.length };
    if (dryRun) {
      return { ...preview, dryRun: true, deleted: false };
    }
    if (params.expectedTitle === undefined) {
      throw new Error('expectedTitle is required to delete (copy it from the dry-run preview)');
    }
    if (params.expectedTitle !== title) {
      throw new Error(`expectedTitle does not match: current title is "${title}"`);
    }
    await rem.remove();
    return { ...preview, dryRun: false, deleted: true };
  }

  // ------------------------------------------------------------------ cards

  async getCards(params: GetCardsParams) {
    const rem = await this.findRem(requireString(params.remId, 'remId'));
    const cards = await rem.getCards();
    const now = Date.now();
    return {
      remId: rem._id,
      title: await this.title(rem),
      cards: cards.map((c) => {
        const history = c.repetitionHistory ?? [];
        const last = history[history.length - 1];
        return {
          cardId: c._id,
          type: c.type,
          nextRepetitionTime: c.nextRepetitionTime,
          isDue: c.nextRepetitionTime !== undefined && c.nextRepetitionTime <= now,
          lastRepetitionTime: c.lastRepetitionTime,
          timesWrongInRow: c.timesWrongInRow,
          reviewCount: history.length,
          lastScore: last?.score,
          ...(params.includeHistory ? { history } : {}),
        };
      }),
    };
  }

  // ------------------------------------------------------------------ helpers

  private requireWrite(): void {
    if (!this.getSettings().acceptWriteOperations) {
      throw new Error('Write operations are disabled in Automation Bridge settings');
    }
  }

  private async findRem(remId: string): Promise<PluginRemLike> {
    const rem = await this.plugin.rem.findOne(remId);
    if (!rem) {
      throw new Error(`Note not found: ${remId}`);
    }
    return rem;
  }

  private async findUploadsFolder(remId?: string): Promise<PluginRemLike | undefined> {
    if (remId) {
      return await this.findRem(remId);
    }
    const hits = await this.plugin.search.search([UPLOADED_FILES_FOLDER_TITLE], undefined, {
      numResults: 20,
    });
    for (const rem of hits) {
      if ((await this.title(rem)).trim() === UPLOADED_FILES_FOLDER_TITLE && !rem.parent) {
        return rem;
      }
    }
    return undefined;
  }

  private async title(rem: PluginRemLike): Promise<string> {
    return await this.plugin.richText.toString(rem.text ?? []);
  }

  private async sourceIds(rem: PluginRemLike): Promise<string[]> {
    return (await rem.getSources()).map((s) => s._id);
  }

  private async uploadedFileUrl(rem: PluginRemLike): Promise<string | undefined> {
    try {
      if (!(await rem.hasPowerup(BuiltInPowerupCodes.UploadedFile))) return undefined;
      return (await rem.getPowerupProperty(BuiltInPowerupCodes.UploadedFile, 'URL')) || undefined;
    } catch {
      return undefined;
    }
  }
}

function requireString(value: unknown, name: string): string {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(`${name} must be a non-empty string`);
  }
  return value;
}

function requireBoolean(value: unknown, name: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`${name} must be a boolean`);
  }
  return value;
}
