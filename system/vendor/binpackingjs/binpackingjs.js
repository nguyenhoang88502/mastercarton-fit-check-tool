/*
 * Vendored from https://github.com/olragon/binpackingjs (npm: binpackingjs@4.1.0), MIT License.
 * Source: dist/cjs/index.js, fetched via https://cdn.jsdelivr.net/npm/binpackingjs@4.1.0/dist/cjs/index.js
 * Wrapped below in a CommonJS shim so it can load as a plain <script> with no bundler.
 * Exposes window.BinPackingJS = { pack2D, pack3D, Packer2D, Packer3D, RotationType, Axis, ALL_ROTATIONS, ... }
 * See ./LICENSE for the original license text.
 */
(function (global) {
  var module = { exports: {} };
  var exports = module.exports;
  var __defProp = Object.defineProperty;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  function __accessProp(key) {
    return this[key];
  }
  var __toCommonJS = (from) => {
    var entry = (__moduleCache ??= new WeakMap).get(from), desc;
    if (entry)
      return entry;
    entry = __defProp({}, "__esModule", { value: true });
    if (from && typeof from === "object" || typeof from === "function") {
      for (var key of __getOwnPropNames(from))
        if (!__hasOwnProp.call(entry, key))
          __defProp(entry, key, {
            get: __accessProp.bind(from, key),
            enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable
          });
    }
    __moduleCache.set(from, entry);
    return entry;
  };
  var __moduleCache;
  var __returnValue = (v) => v;
  function __exportSetter(name, newValue) {
    this[name] = __returnValue.bind(null, newValue);
  }
  var __export = (target, all) => {
    for (var name in all)
      __defProp(target, name, {
        get: all[name],
        enumerable: true,
        configurable: true,
        set: __exportSetter.bind(all, name)
      });
  };

  // src/2d/index.ts
  var exports_2d = {};
  __export(exports_2d, {
    toOriginal: () => toOriginal,
    pack2D: () => pack2D,
    factoredInteger: () => factoredInteger,
    computeFactor: () => computeFactor,
    Score: () => Score,
    Packer2D: () => Packer2D,
    BottomLeft: () => BottomLeft,
    BestShortSideFit: () => BestShortSideFit,
    BestLongSideFit: () => BestLongSideFit,
    BestAreaFit: () => BestAreaFit
  });
  module.exports = __toCommonJS(exports_2d);

  // src/2d/score.ts
  class Score {
    score1;
    score2;
    static MAX_INT = Number.MAX_SAFE_INTEGER;
    static MAX = new Score(Score.MAX_INT, Score.MAX_INT);
    constructor(score1, score2) {
      this.score1 = score1;
      this.score2 = score2;
    }
    isBlank() {
      return this.score1 === Score.MAX_INT;
    }
    isBetterThan(other) {
      if (this.score1 < other.score1)
        return true;
      if (this.score1 === other.score1 && this.score2 < other.score2)
        return true;
      return false;
    }
    subtract(delta) {
      return new Score(this.score1 - delta, this.score2 - delta);
    }
  }

  // src/2d/heuristics/base.ts
  class BaseHeuristic {
    findPosition(boxWidth, boxHeight, freeRects, constrainRotation) {
      let bestScore = Score.MAX;
      let bestPlacement = null;
      for (const freeRect of freeRects) {
        const normal = this.tryFit(freeRect, boxWidth, boxHeight);
        if (normal && normal.score.isBetterThan(bestScore)) {
          bestScore = normal.score;
          bestPlacement = normal;
        }
        if (!constrainRotation) {
          const rotated = this.tryFit(freeRect, boxHeight, boxWidth);
          if (rotated && rotated.score.isBetterThan(bestScore)) {
            bestScore = rotated.score;
            bestPlacement = rotated;
          }
        }
      }
      return bestPlacement;
    }
    tryFit(freeRect, rectWidth, rectHeight) {
      if (rectWidth > freeRect.width || rectHeight > freeRect.height) {
        return null;
      }
      const score = this.calculateScore(freeRect, rectWidth, rectHeight);
      return {
        x: freeRect.x,
        y: freeRect.y,
        width: rectWidth,
        height: rectHeight,
        score
      };
    }
  }

  // src/2d/heuristics/best-short-side-fit.ts
  class BestShortSideFit extends BaseHeuristic {
    calculateScore(freeRect, rectWidth, rectHeight) {
      const leftOverHoriz = Math.abs(freeRect.width - rectWidth);
      const leftOverVert = Math.abs(freeRect.height - rectHeight);
      const shortSide = Math.min(leftOverHoriz, leftOverVert);
      const longSide = Math.max(leftOverHoriz, leftOverVert);
      return new Score(shortSide, longSide);
    }
  }

  // src/2d/free-rect-manager.ts
  class FreeRectManager {
    rects;
    constructor(width, height) {
      this.rects = [{ x: 0, y: 0, width, height }];
    }
    getRects() {
      return this.rects;
    }
    clone() {
      const copy = Object.create(FreeRectManager.prototype);
      copy.rects = this.rects.map((r) => ({ ...r }));
      return copy;
    }
    insert(placed) {
      let numToProcess = this.rects.length;
      let i = 0;
      while (i < numToProcess) {
        if (this.splitFreeNode(this.rects[i], placed)) {
          this.rects.splice(i, 1);
          numToProcess--;
        } else {
          i++;
        }
      }
      this.pruneFreeList();
    }
    splitFreeNode(freeNode, usedNode) {
      if (usedNode.x >= freeNode.x + freeNode.width || usedNode.x + usedNode.width <= freeNode.x || usedNode.y >= freeNode.y + freeNode.height || usedNode.y + usedNode.height <= freeNode.y) {
        return false;
      }
      this.trySplitVertically(freeNode, usedNode);
      this.trySplitHorizontally(freeNode, usedNode);
      return true;
    }
    trySplitVertically(freeNode, usedNode) {
      if (usedNode.x < freeNode.x + freeNode.width && usedNode.x + usedNode.width > freeNode.x) {
        if (usedNode.y > freeNode.y && usedNode.y < freeNode.y + freeNode.height) {
          this.rects.push({
            x: freeNode.x,
            y: freeNode.y,
            width: freeNode.width,
            height: usedNode.y - freeNode.y
          });
        }
        if (usedNode.y + usedNode.height < freeNode.y + freeNode.height) {
          this.rects.push({
            x: freeNode.x,
            y: usedNode.y + usedNode.height,
            width: freeNode.width,
            height: freeNode.y + freeNode.height - (usedNode.y + usedNode.height)
          });
        }
      }
    }
    trySplitHorizontally(freeNode, usedNode) {
      if (usedNode.y < freeNode.y + freeNode.height && usedNode.y + usedNode.height > freeNode.y) {
        if (usedNode.x > freeNode.x && usedNode.x < freeNode.x + freeNode.width) {
          this.rects.push({
            x: freeNode.x,
            y: freeNode.y,
            width: usedNode.x - freeNode.x,
            height: freeNode.height
          });
        }
        if (usedNode.x + usedNode.width < freeNode.x + freeNode.width) {
          this.rects.push({
            x: usedNode.x + usedNode.width,
            y: freeNode.y,
            width: freeNode.x + freeNode.width - (usedNode.x + usedNode.width),
            height: freeNode.height
          });
        }
      }
    }
    pruneFreeList() {
      let i = 0;
      while (i < this.rects.length) {
        let j = i + 1;
        let pruned = false;
        while (j < this.rects.length) {
          if (this.isContainedIn(this.rects[i], this.rects[j])) {
            this.rects.splice(i, 1);
            pruned = true;
            break;
          }
          if (this.isContainedIn(this.rects[j], this.rects[i])) {
            this.rects.splice(j, 1);
          } else {
            j++;
          }
        }
        if (!pruned)
          i++;
      }
    }
    isContainedIn(rectA, rectB) {
      return rectA.x >= rectB.x && rectA.y >= rectB.y && rectA.x + rectA.width <= rectB.x + rectB.width && rectA.y + rectA.height <= rectB.y + rectB.height;
    }
  }

  // src/2d/scoreboard-entry.ts
  class ScoreBoardEntry {
    binIndex;
    boxIndex;
    score;
    constructor(binIndex, boxIndex, score) {
      this.binIndex = binIndex;
      this.boxIndex = boxIndex;
      this.score = score;
    }
  }

  // src/2d/scoreboard.ts
  function findBestEntry(bins, boxes, packedFlags) {
    let bestEntry = null;
    let bestScore = Score.MAX;
    for (let bi = 0;bi < bins.length; bi++) {
      const bin = bins[bi];
      for (let boxi = 0;boxi < boxes.length; boxi++) {
        if (packedFlags[boxi])
          continue;
        const box = boxes[boxi];
        if (box.width > bin.width && box.height > bin.width)
          continue;
        if (box.width > bin.height && box.height > bin.height)
          continue;
        const freeRectsCopy = bin.freeRects.clone();
        const placement = bin.heuristic.findPosition(box.width, box.height, freeRectsCopy.getRects(), box.constrainRotation);
        if (placement && placement.score.isBetterThan(bestScore)) {
          bestScore = placement.score;
          bestEntry = new ScoreBoardEntry(bi, boxi, placement.score);
        }
      }
    }
    return bestEntry;
  }

  // src/lib/factor.ts
  function getDecimalPlaces(value) {
    if (!Number.isFinite(value))
      return 0;
    const str = String(value);
    const dotIndex = str.indexOf(".");
    if (dotIndex === -1)
      return 0;
    const expIndex = str.indexOf("e-");
    if (expIndex !== -1) {
      const mantissaDecimals = str.substring(dotIndex + 1, expIndex).length;
      const exp = parseInt(str.substring(expIndex + 2), 10);
      return mantissaDecimals + exp;
    }
    return str.length - dotIndex - 1;
  }
  var MAX_DECIMAL_PLACES = 10;
  var DEFAULT_FACTOR = 1e5;
  function computeFactor(values) {
    let maxDecimals = 0;
    for (const v of values) {
      const d = getDecimalPlaces(v);
      if (d > maxDecimals)
        maxDecimals = d;
    }
    if (maxDecimals === 0)
      return 1;
    if (maxDecimals > MAX_DECIMAL_PLACES)
      maxDecimals = MAX_DECIMAL_PLACES;
    return Math.pow(10, maxDecimals);
  }
  function factoredInteger(value, factor = DEFAULT_FACTOR) {
    return Math.round(value * factor);
  }
  function toOriginal(value, factor = DEFAULT_FACTOR) {
    if (factor === 1)
      return value;
    const decimalPlaces = Math.round(Math.log10(factor));
    return Number((value / factor).toFixed(decimalPlaces));
  }

  // src/2d/packer.ts
  function pack2D(options) {
    const { bins, boxes, heuristic, limit } = options;
    const defaultHeuristic = heuristic ?? new BestShortSideFit;
    const allValues = [];
    for (const bin of bins) {
      allValues.push(bin.width, bin.height);
    }
    for (const box of boxes) {
      allValues.push(box.width, box.height);
    }
    const factor = options.factor ?? computeFactor(allValues);
    const factoredBoxes = boxes.map((box) => {
      const fb = {
        width: factoredInteger(box.width, factor),
        height: factoredInteger(box.height, factor),
        ...box.constrainRotation !== undefined && { constrainRotation: box.constrainRotation }
      };
      return fb;
    });
    const workingBins = bins.map((bin) => {
      const w = factoredInteger(bin.width, factor);
      const h = factoredInteger(bin.height, factor);
      return {
        source: bin,
        width: w,
        height: h,
        freeRects: new FreeRectManager(w, h),
        heuristic: defaultHeuristic,
        packedBoxes: []
      };
    });
    const packedFlags = new Array(boxes.length).fill(false);
    let packedCount = 0;
    const maxPack = limit ?? boxes.length;
    while (packedCount < maxPack) {
      const entry = findBestEntry(workingBins, factoredBoxes, packedFlags);
      if (!entry)
        break;
      const bin = workingBins[entry.binIndex];
      const fBox = factoredBoxes[entry.boxIndex];
      const placement = bin.heuristic.findPosition(fBox.width, fBox.height, bin.freeRects.getRects(), fBox.constrainRotation);
      if (!placement)
        break;
      bin.freeRects.insert(placement);
      const sourceBox = boxes[entry.boxIndex];
      const rotated = placement.width !== fBox.width || placement.height !== fBox.height;
      const defactor = (v) => toOriginal(v, factor);
      bin.packedBoxes.push({
        width: rotated ? sourceBox.height : sourceBox.width,
        height: rotated ? sourceBox.width : sourceBox.height,
        x: defactor(placement.x),
        y: defactor(placement.y),
        rotated,
        sourceBox
      });
      packedFlags[entry.boxIndex] = true;
      packedCount++;
    }
    const packedBins = workingBins.map((wb) => {
      const binArea = wb.source.width * wb.source.height;
      const boxesArea = wb.packedBoxes.reduce((sum, b) => sum + b.width * b.height, 0);
      return {
        width: wb.source.width,
        height: wb.source.height,
        boxes: wb.packedBoxes,
        efficiency: binArea > 0 ? boxesArea * 100 / binArea : 0
      };
    });
    const unpackedBoxes = boxes.filter((_, i) => !packedFlags[i]);
    return { packedBins, unpackedBoxes };
  }

  class Packer2D {
    bins;
    heuristic;
    constructor(bins = [], heuristic) {
      this.bins = bins;
      this.heuristic = heuristic ?? new BestShortSideFit;
    }
    pack(boxes, options) {
      return pack2D({
        bins: this.bins,
        boxes,
        heuristic: this.heuristic,
        limit: options?.limit
      });
    }
  }
  // src/2d/heuristics/best-area-fit.ts
  class BestAreaFit extends BaseHeuristic {
    calculateScore(freeRect, rectWidth, rectHeight) {
      const areaFit = freeRect.width * freeRect.height - rectWidth * rectHeight;
      const leftOverHoriz = Math.abs(freeRect.width - rectWidth);
      const leftOverVert = Math.abs(freeRect.height - rectHeight);
      const shortSide = Math.min(leftOverHoriz, leftOverVert);
      return new Score(areaFit, shortSide);
    }
  }
  // src/2d/heuristics/best-long-side-fit.ts
  class BestLongSideFit extends BaseHeuristic {
    calculateScore(freeRect, rectWidth, rectHeight) {
      const leftOverHoriz = Math.abs(freeRect.width - rectWidth);
      const leftOverVert = Math.abs(freeRect.height - rectHeight);
      const longSide = Math.max(leftOverHoriz, leftOverVert);
      const shortSide = Math.min(leftOverHoriz, leftOverVert);
      return new Score(longSide, shortSide);
    }
  }
  // src/2d/heuristics/bottom-left.ts
  class BottomLeft extends BaseHeuristic {
    calculateScore(freeRect, _rectWidth, rectHeight) {
      return new Score(freeRect.y + rectHeight, freeRect.x);
    }
  }
  // src/3d/index.ts
  var exports_3d = {};
  __export(exports_3d, {
    toOriginal: () => toOriginal,
    scoreRotation: () => scoreRotation,
    rectIntersect: () => rectIntersect,
    pack3D: () => pack3D,
    itemsIntersect: () => itemsIntersect,
    getDimension: () => getDimension,
    factoredInteger: () => factoredInteger,
    computeFactor: () => computeFactor,
    START_POSITION: () => START_POSITION,
    RotationType: () => RotationType,
    Packer3D: () => Packer3D,
    Axis: () => Axis,
    ALL_ROTATIONS: () => ALL_ROTATIONS
  });
  module.exports = __toCommonJS(exports_3d);

  // src/3d/types.ts
  var RotationType;
  ((RotationType2) => {
    RotationType2[RotationType2["WHD"] = 0] = "WHD";
    RotationType2[RotationType2["HWD"] = 1] = "HWD";
    RotationType2[RotationType2["HDW"] = 2] = "HDW";
    RotationType2[RotationType2["DHW"] = 3] = "DHW";
    RotationType2[RotationType2["DWH"] = 4] = "DWH";
    RotationType2[RotationType2["WDH"] = 5] = "WDH";
  })(RotationType ||= {});
  var ALL_ROTATIONS = [
    0 /* WHD */,
    1 /* HWD */,
    2 /* HDW */,
    3 /* DHW */,
    4 /* DWH */,
    5 /* WDH */
  ];
  var Axis;
  ((Axis2) => {
    Axis2[Axis2["Width"] = 0] = "Width";
    Axis2[Axis2["Height"] = 1] = "Height";
    Axis2[Axis2["Depth"] = 2] = "Depth";
  })(Axis ||= {});
  var START_POSITION = [0, 0, 0];

  // src/3d/item.ts
  function getDimension(width, height, depth, rotation) {
    switch (rotation) {
      case 0 /* WHD */:
        return [width, height, depth];
      case 1 /* HWD */:
        return [height, width, depth];
      case 2 /* HDW */:
        return [height, depth, width];
      case 3 /* DHW */:
        return [depth, height, width];
      case 4 /* DWH */:
        return [depth, width, height];
      case 5 /* WDH */:
        return [width, depth, height];
    }
  }
  function rectIntersect(i1Pos, i1Dim, i2Pos, i2Dim, x, y) {
    const cx1 = i1Pos[x] + i1Dim[x] / 2;
    const cy1 = i1Pos[y] + i1Dim[y] / 2;
    const cx2 = i2Pos[x] + i2Dim[x] / 2;
    const cy2 = i2Pos[y] + i2Dim[y] / 2;
    const ix = Math.max(cx1, cx2) - Math.min(cx1, cx2);
    const iy = Math.max(cy1, cy2) - Math.min(cy1, cy2);
    return ix < (i1Dim[x] + i2Dim[x]) / 2 && iy < (i1Dim[y] + i2Dim[y]) / 2;
  }
  function itemsIntersect(pos1, dim1, pos2, dim2) {
    return rectIntersect(pos1, dim1, pos2, dim2, 0 /* Width */, 1 /* Height */) && rectIntersect(pos1, dim1, pos2, dim2, 1 /* Height */, 2 /* Depth */) && rectIntersect(pos1, dim1, pos2, dim2, 0 /* Width */, 2 /* Depth */);
  }
  function normalizeItem(item, factor) {
    return {
      width: factoredInteger(item.width, factor),
      height: factoredInteger(item.height, factor),
      depth: factoredInteger(item.depth, factor),
      weight: factoredInteger(item.weight, factor)
    };
  }
  function getVolume(width, height, depth) {
    return width * height * depth;
  }

  // src/lib/log.ts
  var enabled = false;
  function createLogger(prefix) {
    return (...args) => {
      if (enabled) {
        console.debug(prefix, ...args);
      }
    };
  }

  // src/3d/bin.ts
  var log = createLogger("3D:");
  function normalizeBin(bin, factor) {
    return {
      name: bin.name,
      width: factoredInteger(bin.width, factor),
      height: factoredInteger(bin.height, factor),
      depth: factoredInteger(bin.depth, factor),
      maxWeight: factoredInteger(bin.maxWeight, factor)
    };
  }
  function binVolume(bin) {
    return bin.width * bin.height * bin.depth;
  }
  function scoreRotation(bin, itemWidth, itemHeight, itemDepth, rotation) {
    const d = getDimension(itemWidth, itemHeight, itemDepth, rotation);
    if (bin.width < d[0] || bin.height < d[1] || bin.depth < d[2]) {
      return 0;
    }
    const widthEfficiency = Math.floor(bin.width / d[0]) * d[0] / bin.width;
    const heightEfficiency = Math.floor(bin.height / d[1]) * d[1] / bin.height;
    const depthEfficiency = Math.floor(bin.depth / d[2]) * d[2] / bin.depth;
    return widthEfficiency * heightEfficiency * depthEfficiency;
  }
  function getBestRotationOrder(bin, itemWidth, itemHeight, itemDepth, allowedRotations) {
    const scores = allowedRotations.map((r) => ({
      rotation: r,
      score: scoreRotation(bin, itemWidth, itemHeight, itemDepth, r)
    }));
    return scores.sort((a, b) => b.score - a.score).map((s) => s.rotation);
  }

  class MutableBin3D {
    name;
    width;
    height;
    depth;
    maxWeight;
    items = [];
    constructor(bin) {
      this.name = bin.name;
      this.width = bin.width;
      this.height = bin.height;
      this.depth = bin.depth;
      this.maxWeight = bin.maxWeight;
    }
    getVolume() {
      return this.width * this.height * this.depth;
    }
    getPackedWeight() {
      return this.items.reduce((w, item) => w + item.weight, 0);
    }
    weighItem(weight) {
      if (this.maxWeight === 0)
        return false;
      return this.maxWeight === Infinity || weight + this.getPackedWeight() <= this.maxWeight;
    }
    putItem(sourceItem, itemWidth, itemHeight, itemDepth, itemWeight, position, allowedRotations) {
      const rotations = getBestRotationOrder(this, itemWidth, itemHeight, itemDepth, allowedRotations);
      for (const rotation of rotations) {
        const d = getDimension(itemWidth, itemHeight, itemDepth, rotation);
        if (this.width < position[0] + d[0] || this.height < position[1] + d[1] || this.depth < position[2] + d[2]) {
          continue;
        }
        let fits = true;
        for (const existing of this.items) {
          if (itemsIntersect(position, d, existing.position, existing.dimension)) {
            fits = false;
            break;
          }
        }
        if (fits) {
          const packed = {
            name: sourceItem.name,
            width: itemWidth,
            height: itemHeight,
            depth: itemDepth,
            weight: itemWeight,
            position: [position[0], position[1], position[2]],
            rotationType: rotation,
            dimension: [d[0], d[1], d[2]],
            sourceItem
          };
          this.items.push(packed);
          log("putItem success", packed.name, "at", packed.position, "dim", packed.dimension);
          return packed;
        }
      }
      return null;
    }
  }

  // src/3d/packer.ts
  function normalize(item, factor) {
    const n = normalizeItem(item, factor);
    return {
      source: item,
      width: n.width,
      height: n.height,
      depth: n.depth,
      weight: n.weight,
      allowedRotations: item.allowedRotations ?? ALL_ROTATIONS
    };
  }
  function findFittedBin(bins, item) {
    for (const b of bins) {
      if (!b.weighItem(item.weight))
        continue;
      const startPos = [0, 0, 0];
      const result = b.putItem(item.source, item.width, item.height, item.depth, item.weight, startPos, item.allowedRotations);
      if (result) {
        b.items.pop();
        return b;
      }
    }
    return null;
  }
  function getBiggerBinThan(bins, b, visited) {
    const v = b.getVolume();
    for (const b2 of bins) {
      if (b2.getVolume() > v && !visited.has(b2)) {
        return b2;
      }
    }
    return null;
  }
  function packToBin(bins, b, items, visited = new Set) {
    const firstItem = items[0];
    if (!firstItem)
      return [];
    visited.add(b);
    if (!b.weighItem(firstItem.weight)) {
      const b2 = getBiggerBinThan(bins, b, visited);
      if (b2)
        return packToBin(bins, b2, items, visited);
      return items;
    }
    const startPos = [0, 0, 0];
    const fit = b.putItem(firstItem.source, firstItem.width, firstItem.height, firstItem.depth, firstItem.weight, startPos, firstItem.allowedRotations);
    if (!fit) {
      const b2 = getBiggerBinThan(bins, b, visited);
      if (b2)
        return packToBin(bins, b2, items, visited);
      return items;
    }
    const unpacked = [];
    for (let i = 1;i < items.length; i++) {
      const item = items[i];
      let fitted = false;
      if (b.weighItem(item.weight)) {
        lookup:
          for (let pt = 0;pt < 3; pt++) {
            for (const ib of b.items) {
              let pv;
              const d = ib.dimension;
              switch (pt) {
                case 0 /* Width */:
                  pv = [ib.position[0] + d[0], ib.position[1], ib.position[2]];
                  break;
                case 1 /* Height */:
                  pv = [ib.position[0], ib.position[1] + d[1], ib.position[2]];
                  break;
                case 2 /* Depth */:
                  pv = [ib.position[0], ib.position[1], ib.position[2] + d[2]];
                  break;
                default:
                  continue;
              }
              const result = b.putItem(item.source, item.width, item.height, item.depth, item.weight, pv, item.allowedRotations);
              if (result) {
                fitted = true;
                break lookup;
              }
            }
          }
      }
      if (!fitted) {
        unpacked.push(item);
      }
    }
    return unpacked;
  }
  function pack3D(options) {
    const allValues = [];
    for (const bin of options.bins) {
      allValues.push(bin.width, bin.height, bin.depth, bin.maxWeight);
    }
    for (const item of options.items) {
      allValues.push(item.width, item.height, item.depth, item.weight);
    }
    const factor = options.factor ?? computeFactor(allValues);
    const normalizedBins = options.bins.map((bin) => ({ source: bin, normalized: normalizeBin(bin, factor) })).sort((a, b) => binVolume(a.normalized) - binVolume(b.normalized));
    let items = options.items.map((item) => normalize(item, factor)).sort((a, b) => getVolume(b.width, b.height, b.depth) - getVolume(a.width, a.height, a.depth));
    const mutableBins = normalizedBins.map((b) => new MutableBin3D(b.normalized));
    const unfitItems = [];
    while (items.length > 0) {
      const firstItem = items[0];
      const bin = findFittedBin(mutableBins, firstItem);
      if (!bin) {
        unfitItems.push(firstItem.source);
        items = items.slice(1);
        continue;
      }
      items = packToBin(mutableBins, bin, items);
    }
    const defactor = (v) => toOriginal(v, factor);
    const packedBins = mutableBins.map((mb, i) => {
      const src = normalizedBins[i].source;
      return {
        name: mb.name,
        width: src.width,
        height: src.height,
        depth: src.depth,
        maxWeight: src.maxWeight,
        items: mb.items.map((item) => ({
          ...item,
          width: defactor(item.width),
          height: defactor(item.height),
          depth: defactor(item.depth),
          weight: defactor(item.weight),
          position: item.position.map(defactor),
          dimension: item.dimension.map(defactor)
        })),
        volume: src.width * src.height * src.depth
      };
    });
    return { packedBins, unfitItems };
  }

  class Packer3D {
    bins = [];
    items = [];
    addBin(bin) {
      this.bins.push(bin);
    }
    addItem(item) {
      this.items.push(item);
    }
    pack() {
      return pack3D({ bins: this.bins, items: this.items });
    }
  }
  // src/index.ts
  var exports_src = {};
  __export(exports_src, {
    toOriginal: () => toOriginal,
    scoreRotation: () => scoreRotation,
    rectIntersect: () => rectIntersect,
    pack3D: () => pack3D,
    pack2D: () => pack2D,
    itemsIntersect: () => itemsIntersect,
    getDimension: () => getDimension,
    factoredInteger: () => factoredInteger,
    computeFactor: () => computeFactor,
    Score: () => Score,
    START_POSITION: () => START_POSITION,
    RotationType: () => RotationType,
    Packer3D: () => Packer3D,
    Packer2D: () => Packer2D,
    BottomLeft: () => BottomLeft,
    BestShortSideFit: () => BestShortSideFit,
    BestLongSideFit: () => BestLongSideFit,
    BestAreaFit: () => BestAreaFit,
    Axis: () => Axis,
    ALL_ROTATIONS: () => ALL_ROTATIONS
  });
  module.exports = __toCommonJS(exports_src);

  global.BinPackingJS = module.exports;
})(typeof window !== "undefined" ? window : this);
