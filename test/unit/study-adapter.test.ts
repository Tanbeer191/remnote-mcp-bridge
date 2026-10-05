import { describe, expect, it, vi } from 'vitest';
import { StudyAdapter } from '../../src/api/study-adapter';
import {
  getDefaultAutomationBridgeSettings,
  SETTING_ACCEPT_REPLACE_OPERATION,
  type AutomationBridgeSettings,
} from '../../src/settings';

type MockRem = Record<string, ReturnType<typeof vi.fn> | unknown>;

function makeRem(id: string, title: string, extra: Partial<MockRem> = {}): MockRem {
  return {
    _id: id,
    text: [title],
    parent: undefined,
    setText: vi.fn(),
    setParent: vi.fn(),
    addPowerup: vi.fn(),
    setPowerupProperty: vi.fn(),
    setIsDocument: vi.fn(),
    addSource: vi.fn(),
    removeSource: vi.fn(),
    getSources: vi.fn().mockResolvedValue([]),
    isFolder: vi.fn().mockResolvedValue(false),
    setIsFolder: vi.fn(),
    getDescendants: vi.fn().mockResolvedValue([{}, {}]),
    remove: vi.fn(),
    getCards: vi.fn().mockResolvedValue([]),
    hasPowerup: vi.fn().mockResolvedValue(false),
    getPowerupProperty: vi.fn(),
    ...extra,
  };
}

function setup(overrides: Partial<AutomationBridgeSettings> = {}) {
  const rems = new Map<string, MockRem>();
  const created = makeRem('new-pdf', '');
  const plugin = {
    rem: {
      findOne: vi.fn(async (id: string) => rems.get(id)),
      createRem: vi.fn(async () => created),
    },
    search: { search: vi.fn(async () => [...rems.values()]) },
    richText: {
      toString: vi.fn(async (rt: unknown[]) => (rt ?? []).join('')),
    },
    settings: {
      // Live RemNote setting store; delete reads the replace setting from here.
      getSetting: vi.fn(async (id: string) =>
        id === SETTING_ACCEPT_REPLACE_OPERATION ? settings.acceptReplaceOperation : undefined
      ),
    },
  };
  const settings = { ...getDefaultAutomationBridgeSettings(), ...overrides };
  const adapter = new StudyAdapter(plugin as never, () => settings);
  return { adapter, plugin, rems, created };
}

describe('StudyAdapter', () => {
  describe('attachPdf', () => {
    it('previews by default without creating anything', async () => {
      const { adapter, rems, plugin } = setup();
      rems.set('doc', makeRem('doc', 'GI Lecture 1'));
      rems.set('folder', makeRem('folder', 'Uploaded Files'));

      const result = await adapter.attachPdf({
        documentRemId: 'doc',
        url: 'https://remnote-user-data.s3.amazonaws.com/x.pdf',
        fileName: 'slides.pdf',
      });

      expect(result).toMatchObject({ dryRun: true, created: false, pdfParentRemId: 'folder' });
      expect(plugin.rem.createRem).not.toHaveBeenCalled();
    });

    it('creates an uploaded-file Rem in Uploaded Files and adds it as a source', async () => {
      const { adapter, rems, created } = setup();
      const doc = makeRem('doc', 'GI Lecture 1');
      rems.set('doc', doc);
      rems.set('folder', makeRem('folder', 'Uploaded Files'));
      const url = 'https://remnote-user-data.s3.amazonaws.com/x.pdf';

      const result = await adapter.attachPdf({
        documentRemId: 'doc',
        url,
        fileName: 'slides.pdf',
        dryRun: false,
      });

      expect(result).toMatchObject({ created: true, pdfRemId: 'new-pdf' });
      expect(created.setParent).toHaveBeenCalledWith('folder');
      expect(created.addPowerup).toHaveBeenCalledWith('f');
      expect(created.setPowerupProperty).toHaveBeenCalledWith('f', 'URL', [url]);
      expect(created.setPowerupProperty).toHaveBeenCalledWith('f', 'Name', ['slides.pdf']);
      expect(created.setIsDocument).toHaveBeenCalledWith(true);
      expect(doc.addSource).toHaveBeenCalledWith(created);
    });

    it('rejects non-https URLs and disabled writes', async () => {
      const { adapter, rems } = setup();
      rems.set('doc', makeRem('doc', 'Doc'));
      await expect(
        adapter.attachPdf({ documentRemId: 'doc', url: 'file:///x.pdf', fileName: 'x.pdf' })
      ).rejects.toThrow('https');

      const off = setup({ acceptWriteOperations: false });
      await expect(
        off.adapter.attachPdf({ documentRemId: 'doc', url: 'https://a/x.pdf', fileName: 'x.pdf' })
      ).rejects.toThrow('Write operations are disabled');
    });
  });

  describe('deleteNote', () => {
    it('requires the replace setting', async () => {
      const { adapter, rems } = setup({ acceptReplaceOperation: false });
      rems.set('n', makeRem('n', 'Test'));
      await expect(adapter.deleteNote({ remId: 'n' })).rejects.toThrow('Accept replace operation');
    });

    it('previews, then needs the exact title to delete', async () => {
      const { adapter, rems } = setup({ acceptReplaceOperation: true });
      const rem = makeRem('n', 'MCP pipeline test');
      rems.set('n', rem);

      expect(await adapter.deleteNote({ remId: 'n' })).toMatchObject({
        dryRun: true,
        descendantCount: 2,
      });
      await expect(adapter.deleteNote({ remId: 'n', dryRun: false })).rejects.toThrow(
        'expectedTitle is required'
      );
      await expect(
        adapter.deleteNote({ remId: 'n', dryRun: false, expectedTitle: 'Wrong' })
      ).rejects.toThrow('does not match');
      expect(rem.remove).not.toHaveBeenCalled();

      await adapter.deleteNote({ remId: 'n', dryRun: false, expectedTitle: 'MCP pipeline test' });
      expect(rem.remove).toHaveBeenCalledOnce();
    });
  });

  it('setFolderStatus previews by default', async () => {
    const { adapter, rems } = setup();
    const rem = makeRem('n', 'Blocks');
    rems.set('n', rem);
    expect(await adapter.setFolderStatus({ remId: 'n', isFolder: true })).toMatchObject({
      wouldChange: true,
      changed: false,
    });
    expect(rem.setIsFolder).not.toHaveBeenCalled();
    await adapter.setFolderStatus({ remId: 'n', isFolder: true, dryRun: false });
    expect(rem.setIsFolder).toHaveBeenCalledWith(true);
  });

  it('getCards summarises due state and last score', async () => {
    const { adapter, rems } = setup();
    rems.set(
      'n',
      makeRem('n', 'Card note', {
        getCards: vi.fn().mockResolvedValue([
          {
            _id: 'c1',
            type: 'forward',
            nextRepetitionTime: 1,
            repetitionHistory: [{ date: 1, score: 0 }],
            timesWrongInRow: 1,
          },
        ]),
      })
    );
    const result = await adapter.getCards({ remId: 'n' });
    expect(result.cards[0]).toMatchObject({
      cardId: 'c1',
      isDue: true,
      reviewCount: 1,
      lastScore: 0,
    });
    expect(result.cards[0]).not.toHaveProperty('history');
  });
});
