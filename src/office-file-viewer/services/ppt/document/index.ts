// PPT 文档层入口，按持久化对象关系读取文档、母版和幻灯片。
export { readPptBinaryDocument } from './readDocument';
export {
  readPptHyperlinks,
  readPptInteractiveHyperlink,
  readPptTextHyperlinkRanges,
} from './readHyperlinks';
export type { PptTextHyperlinkRange } from './readHyperlinks';
export { readPptMaster } from './readMaster';
export { readPptSlide } from './readSlide';
export { readPptSlideLists } from './readSlideLists';
