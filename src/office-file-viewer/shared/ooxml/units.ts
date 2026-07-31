/** 一英寸包含的 EMU 数量。 */
export const EMU_PER_INCH = 914400;
/** 标准化渲染坐标中一英寸包含的像素数。 */
export const PX_PER_INCH = 96;

/** 将 Office EMU 长度换算为标准化渲染像素。 */
export function emuToPx(emu: number) {
  return (emu / EMU_PER_INCH) * PX_PER_INCH;
}
