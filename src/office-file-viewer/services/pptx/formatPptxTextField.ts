import { attr, childByLocalName } from '../../shared/ooxml/xml';

/** 解析幻灯片动态字段时可用的页面和时间上下文。 */
export type PptxTextFieldContext = {
  /** 当前幻灯片的一基页码。 */
  slideNumber?: number;
  /** 用于更新日期时间字段的时刻；缺省时读取当前时间。 */
  currentDate?: Date;
};

function formatDate(
  date: Date,
  locale: string | undefined,
  options: Intl.DateTimeFormatOptions,
) {
  try {
    return new Intl.DateTimeFormat(locale, options).format(date);
  } catch {
    return new Intl.DateTimeFormat(undefined, options).format(date);
  }
}

function readDatePart(
  date: Date,
  locale: string | undefined,
  options: Intl.DateTimeFormatOptions,
) {
  return formatDate(date, locale, options);
}

function readFieldLocale(field: Element) {
  return (
    attr(childByLocalName(field, 'rPr'), 'lang') ??
    attr(childByLocalName(field, 'endParaRPr'), 'lang')
  );
}

/** 按 PresentationML 保留字段类型和文字语言更新动态日期时间。 */
export function formatPptxTextField(
  field: Element,
  fallback: string,
  context: PptxTextFieldContext = {},
) {
  const type = attr(field, 'type')?.toLowerCase();
  if (type === 'slidenum') {
    return context.slideNumber === undefined
      ? fallback
      : String(context.slideNumber);
  }
  if (!type?.startsWith('datetime')) return fallback;

  const currentDate = context.currentDate ?? new Date();
  const locale = readFieldLocale(field);
  const usesEastAsianOrder = /^(?:zh|ja|ko)(?:-|$)/i.test(locale ?? '');
  const numericDate = () =>
    formatDate(currentDate, locale, {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    });
  const longDate = () =>
    formatDate(currentDate, locale, {
      year: 'numeric',
      month: 'long',
      day: '2-digit',
    });
  const time = (withSeconds: boolean, hour12: boolean) =>
    formatDate(currentDate, locale, {
      hour: '2-digit',
      minute: '2-digit',
      second: withSeconds ? '2-digit' : undefined,
      hour12,
    });
  const year = readDatePart(currentDate, locale, { year: 'numeric' });
  const shortYear = readDatePart(currentDate, locale, { year: '2-digit' });
  const longMonth = readDatePart(currentDate, locale, { month: 'long' });
  const shortMonth = readDatePart(currentDate, locale, { month: 'short' });
  const day = readDatePart(currentDate, locale, { day: '2-digit' });

  switch (type) {
    case 'datetime':
      return `${numericDate()} ${time(false, false)}`;
    case 'datetime1':
      return numericDate();
    case 'datetime2':
      return formatDate(currentDate, locale, {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: '2-digit',
      });
    case 'datetime3':
      return usesEastAsianOrder ? longDate() : `${day} ${longMonth} ${year}`;
    case 'datetime4':
      return usesEastAsianOrder ? longDate() : `${longMonth} ${day}, ${year}`;
    case 'datetime5':
      return `${day}-${shortMonth}-${shortYear}`;
    case 'datetime6':
      return usesEastAsianOrder
        ? formatDate(currentDate, locale, {
            year: 'numeric',
            month: 'long',
          })
        : `${longMonth} ${shortYear}`;
    case 'datetime7':
      return usesEastAsianOrder
        ? formatDate(currentDate, locale, {
            year: '2-digit',
            month: 'short',
          })
        : `${shortMonth}-${shortYear}`;
    case 'datetime8':
      return `${numericDate()} ${time(false, true)}`;
    case 'datetime9':
      return `${numericDate()} ${time(true, true)}`;
    case 'datetime10':
      return time(false, false);
    case 'datetime11':
      return time(true, false);
    case 'datetime12':
      return time(false, true);
    case 'datetime13':
      return time(true, true);
    default:
      return fallback;
  }
}
