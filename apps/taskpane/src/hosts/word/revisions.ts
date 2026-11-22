// Word 本轮修订（Tracked Changes）层
// begin：记录用户原修订模式与已有修订列表 → 打开 trackAll。end（务必在 finally 调用）：恢复原模式，算出本轮新增修订。
// 本轮修订 = 前后 getTrackedChanges() 列表之差（按 type + text + date 的多重集合），不能只按时间（作者都是同一 Office 用户、时间只到分钟）。
// 注意：处理时按"当前列表 − 开始前列表"循环，见 settle()。
// "撤销本轮" = 拒绝本轮修订；用户原有修订不受影响。修订保存在文件中，面板刷新不丢。

type Key = string;
const keyOf = (t: { type: string; text: string; date: Date | string }) => `${t.type}|${t.text}|${String(t.date)}`;

function multiset(keys: Key[]) {
  const m = new Map<Key, number>();
  for (const k of keys) m.set(k, (m.get(k) ?? 0) + 1);
  return m;
}

export class WordRevisions {
  readonly kind = 'word';
  private originalMode: string | null = null;
  private before = new Map<Key, number>();
  private turnKeys: Key[] = [];
  state: 'open' | 'pending' | 'accepted' | 'rejected' = 'open';

  get isEmpty() {
    return this.state !== 'open' && this.turnKeys.length === 0;
  }
  // 与 PptChangeset 对齐：已处理（接受或拒绝）的轮次不再可撤销
  get undone() {
    return this.state === 'accepted' || this.state === 'rejected';
  }
  get count() {
    return this.turnKeys.length;
  }

  private async list(ctx: Word.RequestContext) {
    const tc = ctx.document.body.getTrackedChanges();
    tc.load('items/type,items/text,items/date');
    await ctx.sync();
    return tc;
  }

  async begin(): Promise<void> {
    await Word.run(async (ctx) => {
      ctx.document.load('changeTrackingMode');
      const tc = await this.list(ctx);
      this.originalMode = ctx.document.changeTrackingMode;
      this.before = multiset(tc.items.map(keyOf));
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await ctx.sync();
    });
  }

  async end(): Promise<void> {
    await Word.run(async (ctx) => {
      if (this.originalMode) ctx.document.changeTrackingMode = this.originalMode as Word.ChangeTrackingMode;
      const tc = await this.list(ctx);
      const remain = new Map(this.before);
      this.turnKeys = [];
      for (const t of tc.items) {
        const k = keyOf(t);
        const n = remain.get(k) ?? 0;
        if (n > 0) remain.set(k, n - 1);
        else this.turnKeys.push(k);
      }
    });
    this.state = 'pending';
  }

  // 实测（Mac Word 16.112）：一次"替换"在列表里先只出现 Added 条目，配对的 Deleted 条目要等 Added 处理后才出现。
  // 所以循环处理"当前列表 − 本轮开始前已有的修订"，直到为空；用户原有修订始终排除在外。
  private async settle(action: 'accept' | 'reject') {
    await Word.run(async (ctx) => {
      for (let round = 0; round < 5; round++) {
        const tc = await this.list(ctx);
        const remainBefore = new Map(this.before);
        const extra: Word.TrackedChange[] = [];
        for (const t of tc.items) {
          const k = keyOf(t);
          const n = remainBefore.get(k) ?? 0;
          if (n > 0) remainBefore.set(k, n - 1);
          else extra.push(t);
        }
        if (!extra.length) break;
        for (const t of extra) {
          if (action === 'accept') t.accept();
          else t.reject();
        }
        await ctx.sync();
      }
    });
    this.state = action === 'accept' ? 'accepted' : 'rejected';
  }

  acceptTurn() {
    return this.settle('accept');
  }
  rejectTurn() {
    return this.settle('reject');
  }
  // 与 PptChangeset 同名，供"撤销"统一调用
  undo() {
    return this.rejectTurn();
  }
}
