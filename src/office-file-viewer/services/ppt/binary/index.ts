// PPT 二进制基础入口，导出记录读取器、时间片和字符串解码工具。
export { PPT_RECORD } from './constants';
export { PptRecordReader, createPptTimeSlice } from './PptRecordReader';
export { readPptByteString, readPptUnicodeString } from './readStrings';
