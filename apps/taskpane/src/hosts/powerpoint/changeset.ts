// PPT "撤销本轮"快照层（Changeset）
// 写前：把将被修改的幻灯片用 exportAsBase64 导出（M0 已验证无损）；记录本轮新建的幻灯片。
// 撤销：删除新建页 → 按原位置把快照插回并删除被改过的页。插回后幻灯片 ID 会变化。

interface Snapshot {
  base64: string;
  index: number; // 本轮开始时的位置（0 起）
}

export interface ChangesetSummary {
  modified: number;
  created: number;
  modifiedIndexes: number[]; // 1 起，供界面显示
}

export type InsertFormatting = 'KeepSourceFormatting' | 'UseDestinationTheme';
const FORMATTING_KEY = 'keynib.undoFormatting';

// 还原时插回快照所用的格式选项；由诊断页的"撤销自测"决定哪种不会产生重复母版
export function getUndoFormatting(): InsertFormatting {
  try {
    const v = localStorage.getItem(FORMATTING_KEY);
    if (v === 'UseDestinationTheme' || v === 'KeepSourceFormatting') return v;
  } catch {
    /* 忽略 */
  }
  return 'KeepSourceFormatting';
}

export function setUndoFormatting(v: InsertFormatting) {
  try {
    localStorage.setItem(FORMATTING_KEY, v);
  } catch {
    /* 忽略 */
  }
}

export class PptChangeset {
  private formatting: InsertFormatting;
  constructor(formatting: InsertFormatting = getUndoFormatting()) {
    this.formatting = formatting;
  }

  private createdIds = new Set<string>();
  private snapshots = new Map<string, Snapshot>();
  private originalOrder: string[] | null = null;
  undone = false;

  get isEmpty() {
    return this.createdIds.size === 0 && this.snapshots.size === 0;
  }

  summary(): ChangesetSummary {
    return {
      modified: this.snapshots.size,
      created: this.createdIds.size,
      modifiedIndexes: [...this.snapshots.values()].map((s) => s.index + 1).sort((a, b) => a - b),
    };
  }

  noteCreated(slideId: string) {
    this.createdIds.add(slideId);
  }

  // slideIndexes 为当前（1 起）序号
  async beforeWrite(targets: number[] | 'all'): Promise<void> {
    // 只读导出：网页版偶发暂时性 GeneralException，重试两次
    for (let i = 0; ; i++) {
      try {
        return await this.snapshotNow(targets);
      } catch (e: any) {
        if (i >= 2 || e?.code !== 'GeneralException') throw e;
        await new Promise((r) => setTimeout(r, 500 * (i + 1)));
      }
    }
  }

  private async snapshotNow(targets: number[] | 'all'): Promise<void> {
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      await ctx.sync();
      const slideIndexes = targets === 'all' ? slides.items.map((_, i) => i + 1) : targets;
      if (!this.originalOrder) this.originalOrder = slides.items.map((s) => s.id);
      const exports: { id: string; result: OfficeExtension.ClientResult<string> }[] = [];
      for (const n of slideIndexes) {
        const id = slides.items[n - 1]?.id;
        if (!id) continue;
        if (this.createdIds.has(id) || this.snapshots.has(id)) continue;
        if (!this.originalOrder.includes(id)) continue; // 本轮之前就不存在，无需快照
        exports.push({ id, result: slides.getItemAt(n - 1).exportAsBase64() });
      }
      if (!exports.length) return;
      await ctx.sync();
      for (const e of exports) {
        this.snapshots.set(e.id, { base64: e.result.value, index: this.originalOrder.indexOf(e.id) });
      }
    });
  }

  // 新建页之前也需要记录原顺序，保证撤销时位置正确
  async ensureOrder(): Promise<void> {
    if (this.originalOrder) return;
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      await ctx.sync();
      this.originalOrder = slides.items.map((s) => s.id);
    });
  }

  // 网页版实测：按 ID 查找刚新建的页（getItem / getItemOrNullObject）会暂时失败。
  // 因此一律先读出全部 ID 算出位置，再按序号（getItemAt）操作。
  async undo(): Promise<void> {
    if (this.undone) return;
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      const indexOf = async (id: string) => {
        slides.load('items/id');
        await ctx.sync();
        return slides.items.findIndex((s) => s.id === id);
      };

      // 1. 删除本轮新建的页。网页版实测：新建页先拿到临时 ID，之后会换成正式 ID，
      //    所以不能靠新建时记下的 ID 去找；改为"本轮开始时不存在的页 = 本轮新建"（已有页的 ID 稳定）。
      slides.load('items/id');
      await ctx.sync();
      const original = this.originalOrder ? new Set(this.originalOrder) : null;
      const newIdx = slides.items.map((x, i) => (original ? !original.has(x.id) : this.createdIds.has(x.id)) ? i : -1).filter((i) => i >= 0);
      // 保护：判定为新建的页数必须等于本轮实际新建数，否则（如上一轮的新页 ID 尚未稳定）宁可不删，避免误删用户的页
      if (newIdx.length !== this.createdIds.size) {
        throw new Error(
          `无法可靠识别本轮新建的幻灯片（识别到 ${newIdx.length} 页，本轮新建 ${this.createdIds.size} 页），为避免误删已停止撤销，请手动检查`,
        );
      }
      for (const i of [...newIdx].reverse()) slides.getItemAt(i).delete();
      await ctx.sync();

      // 2. 按原位置从前往后还原被修改 / 删除 / 移动过的页
      const restore = [...this.snapshots.entries()].sort((a, b) => a[1].index - b[1].index);
      for (const [id, snap] of restore) {
        const i = await indexOf(id);
        if (i >= 0) {
          slides.getItemAt(i).delete();
          await ctx.sync();
        }
        slides.load('items/id');
        await ctx.sync();
        const prev = snap.index > 0 ? slides.items[Math.min(snap.index, slides.items.length) - 1] : undefined;
        ctx.presentation.insertSlidesFromBase64(snap.base64, {
          formatting: this.formatting,
          ...(prev ? { targetSlideId: prev.id } : {}),
        });
        await ctx.sync();
      }
    });
    this.undone = true;
  }
}
