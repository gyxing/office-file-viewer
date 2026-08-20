import {
  getPresentationMediaMimeType,
  type PresentationMediaKind,
} from '../presentation/mediaTypes';
import { PPT_RECORD } from './binary/constants';
import { readPptUnicodeString } from './binary/readStrings';
import { walkPptRecords } from './binary/walkPptRecords';
import {
  createPptResourceId,
  registerPptResource,
  type PptParseContext,
  type PptRecord,
} from './types';

type EmbeddedSound = Readonly<{
  source: string;
  fileName?: string;
  mimeType: string;
}>;

function collectRecords(root: PptRecord) {
  const records: PptRecord[] = [];
  walkPptRecords(root, (record) => records.push(record));
  return records;
}

function readStrings(records: readonly PptRecord[]) {
  return records
    .filter((record) => record.type === PPT_RECORD.C_STRING)
    .map((record) => ({
      instance: record.instance,
      value: readPptUnicodeString(record.data).trim(),
    }))
    .filter((item) => item.value);
}

function readEmbeddedSounds(
  records: readonly PptRecord[],
  context: PptParseContext,
) {
  const sounds = new Map<number, EmbeddedSound>();
  records
    .filter((record) => record.type === PPT_RECORD.SOUND)
    .forEach((record, fallbackIndex) => {
      const nested = collectRecords(record);
      const strings = readStrings(nested);
      const explicitId = strings
        .filter((item) => item.instance >= 2)
        .map((item) => Number(item.value))
        .find((value) => Number.isInteger(value) && value >= 0);
      const anyId = strings
        .map((item) => Number(item.value))
        .find((value) => Number.isInteger(value) && value >= 0);
      const soundId = explicitId ?? anyId ?? fallbackIndex;
      const fileName = strings.find(
        (item) => !Number.isFinite(Number(item.value)),
      )?.value;
      const data = nested.find(
        (item) => item.type === PPT_RECORD.SOUND_DATA_BLOB,
      );
      if (!data?.data.length) return;
      const mimeType = fileName
        ? getPresentationMediaMimeType(fileName)
        : 'audio/wav';
      const id = createPptResourceId(context, 'media');
      const buffer = data.data.buffer.slice(
        data.data.byteOffset,
        data.data.byteOffset + data.data.byteLength,
      ) as ArrayBuffer;
      sounds.set(soundId, {
        source: registerPptResource(context, {
          id,
          encoding: 'binary',
          mimeType,
          buffer,
        }),
        fileName,
        mimeType,
      });
    });
  return sounds;
}

function mediaKind(recordType: number): PresentationMediaKind {
  return (
    [
      PPT_RECORD.EXTERNAL_VIDEO,
      PPT_RECORD.EXTERNAL_AVI_MOVIE,
      PPT_RECORD.EXTERNAL_MCI_MOVIE,
    ] as number[]
  ).includes(recordType)
    ? 'video'
    : 'audio';
}

function readExternalMedia(
  record: PptRecord,
  sounds: ReadonlyMap<number, EmbeddedSound>,
  context: PptParseContext,
) {
  const nested = collectRecords(record);
  const mediaAtom = nested.find(
    (item) => item.type === PPT_RECORD.EXTERNAL_MEDIA_ATOM,
  );
  if (!mediaAtom || mediaAtom.data.length < 6) return;
  const mediaView = new DataView(
    mediaAtom.data.buffer,
    mediaAtom.data.byteOffset,
    mediaAtom.data.byteLength,
  );
  const objectId = mediaView.getUint32(0, true);
  const flags = mediaView.getUint16(4, true);
  const embeddedAtom = nested.find(
    (item) => item.type === PPT_RECORD.EXTERNAL_WAV_AUDIO_EMBEDDED_ATOM,
  );
  const soundId =
    embeddedAtom && embeddedAtom.data.length >= 4
      ? new DataView(
          embeddedAtom.data.buffer,
          embeddedAtom.data.byteOffset,
          embeddedAtom.data.byteLength,
        ).getUint32(0, true)
      : undefined;
  const embedded = soundId === undefined ? undefined : sounds.get(soundId);
  if (embedded) {
    context.presentationMedia.set(objectId, {
      kind: 'audio',
      sourceKind: 'embedded',
      source: embedded.source,
      mimeType: embedded.mimeType,
      fileName: embedded.fileName,
      loop: Boolean(flags & 0x0001),
    });
    return;
  }

  const target = readStrings(nested)
    .map((item) => item.value)
    .find(
      (value) =>
        !Number.isFinite(Number(value)) &&
        !value.startsWith('___') &&
        /[:\\/]|\.[a-z0-9]{2,5}$/i.test(value),
    );
  if (!target) return;
  context.presentationMedia.set(objectId, {
    kind: mediaKind(record.type),
    sourceKind: 'external',
    source: target,
    mimeType: getPresentationMediaMimeType(target),
    fileName: target.split(/[\\/]/).pop(),
    loop: Boolean(flags & 0x0001),
  });
}

/** 读取 PPT 文档级音视频对象表并登记内嵌声音资源。 */
export function readPptMedia(
  documentRecord: PptRecord,
  context: PptParseContext,
) {
  const records = collectRecords(documentRecord);
  const sounds = readEmbeddedSounds(records, context);
  const mediaTypes = new Set<number>([
    PPT_RECORD.EXTERNAL_VIDEO,
    PPT_RECORD.EXTERNAL_AVI_MOVIE,
    PPT_RECORD.EXTERNAL_MCI_MOVIE,
    PPT_RECORD.EXTERNAL_MIDI_AUDIO,
    PPT_RECORD.EXTERNAL_CD_AUDIO,
    PPT_RECORD.EXTERNAL_WAV_AUDIO_EMBEDDED,
    PPT_RECORD.EXTERNAL_WAV_AUDIO_LINK,
  ]);
  records
    .filter((record) => mediaTypes.has(record.type))
    .forEach((record) => readExternalMedia(record, sounds, context));
}
