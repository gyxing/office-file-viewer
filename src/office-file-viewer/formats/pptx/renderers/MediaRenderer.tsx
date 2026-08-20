import type { CSSProperties } from 'react';
import React, { memo, useEffect, useMemo, useRef, useState } from 'react';
import { useOfficeFileViewerMessages } from '../../../locale';
import type { OfficeFileViewerPresentationMediaOptions } from '../../../services/presentation/mediaTypes';
import type { MediaElement } from '../../../services/presentation/types';
import {
  useOfficeResourceUrl,
  type OfficeResourceSource,
} from '../../../services/resource-store';
import { DownloadIcon } from '../../../shared/ui/OfficeIcons';

/** 判断外部媒体地址是否属于显式允许的 HTTP(S) 读取范围。 */
function isSafeExternalMediaUrl(value: string) {
  try {
    const base =
      typeof document === 'undefined' ? 'https://localhost/' : document.baseURI;
    const protocol = new URL(value, base).protocol.toLowerCase();
    return protocol === 'http:' || protocol === 'https:';
  } catch {
    return false;
  }
}

/** PPT/PPTX 媒体渲染属性。 */
type MediaRendererProps = {
  /** 当前负责渲染的音视频元素。 */
  element: MediaElement;
  /** 缩略图传 false，避免创建媒体 URL 或发起媒体请求。 */
  interactive: boolean;
  /** 外部媒体和下载操作配置。 */
  options?: false | OfficeFileViewerPresentationMediaOptions;
};

/** 使用浏览器原生控件按需播放演示文稿音视频。 */
function MediaRendererComponent({
  element,
  interactive,
  options,
}: MediaRendererProps) {
  const messages = useOfficeFileViewerMessages();
  const mediaRef = useRef<HTMLMediaElement>(null);
  const [playbackError, setPlaybackError] = useState(false);
  const externalAllowed =
    element.media.sourceKind !== 'external' ||
    (options !== false &&
      options?.allowExternal === true &&
      typeof element.media.source === 'string' &&
      isSafeExternalMediaUrl(element.media.source));
  const mediaSource = useMemo<OfficeResourceSource | undefined>(() => {
    if (!interactive || !externalAllowed || options === false) return undefined;
    return typeof element.media.source === 'string'
      ? { kind: 'url', url: element.media.source }
      : element.media.source;
  }, [element.media.source, externalAllowed, interactive, options]);
  const posterSource = useMemo<OfficeResourceSource | undefined>(
    () =>
      typeof element.posterSrc === 'string'
        ? { kind: 'url', url: element.posterSrc }
        : element.posterSrc,
    [element.posterSrc],
  );
  const mediaResource = useOfficeResourceUrl(mediaSource);
  const posterResource = useOfficeResourceUrl(posterSource);
  const frameStyle = useMemo<CSSProperties>(
    () => ({
      left: element.x,
      top: element.y,
      width: element.width,
      height: element.height,
      transform: [
        element.rotate ? `rotate(${element.rotate}deg)` : '',
        element.flipH ? 'scaleX(-1)' : '',
        element.flipV ? 'scaleY(-1)' : '',
      ]
        .filter(Boolean)
        .join(' '),
      zIndex: element.zIndex,
      backgroundImage: posterResource.url
        ? `url(${posterResource.url})`
        : undefined,
    }),
    [
      element.flipH,
      element.flipV,
      element.height,
      element.rotate,
      element.width,
      element.x,
      element.y,
      element.zIndex,
      posterResource.url,
    ],
  );

  useEffect(
    () => () => {
      mediaRef.current?.pause();
    },
    [],
  );

  const label = messages.presentation.mediaLabel(
    element.media.kind,
    element.alt ?? element.media.fileName,
  );
  const blocked =
    interactive &&
    options !== false &&
    element.media.sourceKind === 'external' &&
    !externalAllowed;
  const showDownload =
    interactive &&
    options !== false &&
    options?.download !== false &&
    Boolean(mediaResource.url);

  return (
    <div
      className={[
        'office-file-pptx-media',
        `office-file-pptx-media--${element.media.kind}`,
        interactive ? undefined : 'office-file-pptx-media--thumbnail',
      ]
        .filter(Boolean)
        .join(' ')}
      style={frameStyle}
      data-office-presentation-element-id={element.id}
      aria-label={label}
    >
      {!interactive ? (
        <span className="office-file-pptx-media__poster-icon" />
      ) : null}
      {blocked ? (
        <div className="office-file-pptx-media__message">
          {messages.presentation.externalMediaBlocked}
        </div>
      ) : null}
      {!blocked && mediaResource.loading ? (
        <div className="office-file-pptx-media__message">
          {messages.presentation.mediaLoading}
        </div>
      ) : null}
      {!blocked && mediaResource.url && element.media.kind === 'video' ? (
        <video
          ref={mediaRef as React.RefObject<HTMLVideoElement>}
          aria-label={label}
          controls
          loop={element.media.loop}
          poster={posterResource.url}
          preload="metadata"
          src={mediaResource.url}
          onError={() => setPlaybackError(true)}
        />
      ) : null}
      {!blocked && mediaResource.url && element.media.kind === 'audio' ? (
        <audio
          ref={mediaRef as React.RefObject<HTMLAudioElement>}
          aria-label={label}
          controls
          loop={element.media.loop}
          preload="metadata"
          src={mediaResource.url}
          onError={() => setPlaybackError(true)}
        />
      ) : null}
      {!blocked && (playbackError || mediaResource.error) ? (
        <div className="office-file-pptx-media__message office-file-pptx-media__message--error">
          {messages.presentation.mediaUnsupported}
        </div>
      ) : null}
      {showDownload ? (
        <a
          className="office-file-pptx-media__download"
          href={mediaResource.url}
          download={element.media.fileName}
          target={
            element.media.sourceKind === 'external' ? '_blank' : undefined
          }
          rel="noopener noreferrer"
          onClick={(event) => event.stopPropagation()}
        >
          <DownloadIcon />
          <span>{messages.presentation.downloadMedia}</span>
        </a>
      ) : null}
    </div>
  );
}

export const MediaRenderer = memo(MediaRendererComponent);
