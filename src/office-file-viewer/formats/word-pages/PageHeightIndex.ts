/** 使用 Fenwick Tree 维护页面高度前缀和，查询与替换均为 O(log N)。 */
export class PageHeightIndex {
  private readonly values: number[] = [];
  private readonly tree: number[] = [0];

  constructor(heights: number[]) {
    this.append(heights);
  }

  get length() {
    return this.values.length;
  }

  append(heights: readonly number[]) {
    heights.forEach((rawHeight) => {
      const height = Math.max(0, rawHeight);
      const treeIndex = this.values.length + 1;
      const coveredStart = treeIndex - (treeIndex & -treeIndex);
      const coveredPrevious =
        this.prefix(treeIndex - 1) - this.prefix(coveredStart);
      this.values.push(height);
      this.tree.push(coveredPrevious + height);
    });
  }

  replace(index: number, rawHeight: number) {
    if (index < 0 || index >= this.values.length) return 0;
    const height = Math.max(0, rawHeight);
    const delta = height - this.values[index];
    if (!delta) return 0;
    this.values[index] = height;
    for (
      let treeIndex = index + 1;
      treeIndex < this.tree.length;
      treeIndex += treeIndex & -treeIndex
    ) {
      this.tree[treeIndex] += delta;
    }
    return delta;
  }

  prefix(indexExclusive: number) {
    let sum = 0;
    for (
      let treeIndex = Math.min(Math.max(0, indexExclusive), this.values.length);
      treeIndex > 0;
      treeIndex -= treeIndex & -treeIndex
    ) {
      sum += this.tree[treeIndex];
    }
    return sum;
  }

  total() {
    return this.prefix(this.values.length);
  }

  /** 返回指定纵向偏移所在的页面索引。 */
  findIndexAtOffset(offset: number) {
    if (!this.values.length) return -1;
    if (offset <= 0) return 0;
    if (offset >= this.total()) return this.values.length - 1;

    let treeIndex = 0;
    let accumulated = 0;
    let step = 1;
    while (step * 2 < this.tree.length) step *= 2;
    for (; step > 0; step = Math.floor(step / 2)) {
      const next = treeIndex + step;
      if (next < this.tree.length && accumulated + this.tree[next] <= offset) {
        treeIndex = next;
        accumulated += this.tree[next];
      }
    }
    return Math.min(treeIndex, this.values.length - 1);
  }
}
