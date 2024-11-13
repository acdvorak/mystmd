import type {
  Footer,
  IImageOptions,
  INumberingOptions,
  ISectionOptions,
  ParagraphChild,
} from 'docx';
import {
  InternalHyperlink,
  SimpleField,
  Bookmark,
  SequentialIdentifier,
  TextRun,
  Document,
  Packer,
  SectionType,
} from 'docx';
import { Buffer } from 'buffer'; // Important for frontend development!
import type { Image as MdastImage } from 'myst-spec';
import type { PageFrontmatter } from 'myst-frontmatter';
import { selectAll } from 'unist-util-select';
import type { IFootnotes, Options } from './types.js';
import type { GenericParent } from 'myst-common';
import imageSize from 'image-size';
import { Resvg } from '@resvg/resvg-js';
import type { ISizeCalculationResult } from 'image-size/dist/types/interface.js';

export function createShortId() {
  return Math.random().toString(36).slice(2);
}

export function createDocFromState(
  state: {
    numbering: INumberingOptions['config'];
    children: ISectionOptions['children'];
    frontmatter: PageFrontmatter;
    footnotes?: IFootnotes;
  },
  footer?: Footer,
  styles?: string,
) {
  const { title, description, keywords } = state.frontmatter;
  const doc = new Document({
    title,
    description,
    keywords: keywords?.join(', '),
    footnotes: state.footnotes,
    numbering: {
      config: state.numbering,
    },
    sections: [
      {
        properties: {
          type: SectionType.CONTINUOUS,
        },
        children: state.children,
        footers: footer ? { default: footer } : undefined,
      },
    ],
    externalStyles: styles,
  });
  return doc;
}

export async function writeDocx(
  doc: Document,
  write: ((buffer: Buffer) => void) | ((buffer: Buffer) => Promise<void>),
) {
  const buffer = await Packer.toBuffer(doc);
  return write(buffer);
}

export type ImageType = IImageOptions['type'];

export const SUPPORTED_IMAGE_TYPES: readonly ImageType[] = ['bmp', 'gif', 'jpg', 'png', 'svg'];

function isSupportedImageType(type: string | undefined): type is ImageType {
  return Boolean(type) && SUPPORTED_IMAGE_TYPES.includes(type as ImageType);
}

export function getImageType(buffer: Buffer): ImageType | undefined {
  const { type } = imageSize.imageSize(buffer);
  return isSupportedImageType(type) ? type : undefined;
}

export function svgToPng(svg: string | Buffer): Uint8Array {
  const resvg = new Resvg(svg, {});
  const pngData = resvg.render();
  const pngBuffer = pngData.asPng();
  return Uint8Array.from(pngBuffer);
}

const DEFAULT_IMAGE_WIDTH = 70;

const DEFAULT_PAGE_WIDTH_PIXELS = 800;

/**
 * In pixels.
 *
 * Slightly narrower than the default document page width of an 8.5x11 inch
 * Word doc.
 */
export const MAX_DOCX_IMAGE_WIDTH = 600;

export function getImageWidth(width?: number | string, maxWidth = MAX_DOCX_IMAGE_WIDTH): number {
  if (typeof width === 'number' && Number.isNaN(width)) {
    // If it is nan, return with the default.
    return getImageWidth(DEFAULT_IMAGE_WIDTH);
  }
  if (typeof width === 'string') {
    if (width.endsWith('%')) {
      return getImageWidth(Number(width.replace('%', '')));
    } else if (width.endsWith('px')) {
      return getImageWidth(Number(width.replace('px', '')) / DEFAULT_PAGE_WIDTH_PIXELS);
    }
    console.log(`Unsupported width value \`${width}\` in getImageWidth()`);
    return getImageWidth(DEFAULT_IMAGE_WIDTH);
  }
  let lineWidth = width ?? DEFAULT_IMAGE_WIDTH;
  if (lineWidth < 1) lineWidth *= 100;
  if (lineWidth > 100) lineWidth = 100;
  return (lineWidth / 100) * maxWidth;
}

export interface ImageSize {
  width: number;
  height: number;
}

async function getImageDimensions(file: Blob | Buffer): Promise<ImageSize> {
  let size: ISizeCalculationResult;
  if (Buffer.isBuffer(file)) {
    size = imageSize.imageSize(file);
  } else {
    const arrayBuffer = await file.arrayBuffer();
    size = imageSize.imageSize(new Uint8Array(arrayBuffer));
  }
  const { width, height } = size;
  return {
    width: width || DEFAULT_IMAGE_WIDTH,
    height: height || DEFAULT_IMAGE_WIDTH,
  };
}

/**
 * For frontend development, fetch images as Blobs, get their dimensions and
 * return options for the docx serializer.
 *
 * @param tree the mdast document
 * @returns options for the serializer
 */
export async function fetchImagesAsBuffers(
  tree: GenericParent,
): Promise<Required<Pick<Options, 'getImageBuffer' | 'getImageDimensions'>>> {
  const images = selectAll('image', tree) as MdastImage[];
  const buffers: Record<string, Buffer> = {};
  const dimensions: Record<string, { width: number; height: number }> = {};
  await Promise.all(
    images.map(async (image) => {
      const response = await fetch(image.url);
      const blob = await response.blob();
      const buffer = await blob.arrayBuffer();
      dimensions[image.url] = await getImageDimensions(blob);
      buffers[image.url] = Buffer.from(buffer);
    }),
  );
  return {
    getImageBuffer(url: string) {
      return buffers[url];
    },
    getImageDimensions(url: string) {
      return dimensions[url];
    },
  };
}

export function createReferenceBookmark(
  id: string,
  kind: 'Equation' | 'Figure' | 'Table',
  before?: string,
  after?: string,
) {
  const textBefore = before ? [new TextRun(before)] : [];
  const textAfter = after ? [new TextRun(after)] : [];
  return new Bookmark({
    id,
    children: [...textBefore, new SequentialIdentifier(kind), ...textAfter],
  });
}

export function createReference(id: string, before?: string, after?: string) {
  const children: ParagraphChild[] = [];
  if (before) children.push(new TextRun(before));
  children.push(new SimpleField(`REF ${id} \\h`));
  if (after) children.push(new TextRun(after));
  const ref = new InternalHyperlink({ anchor: id, children });
  return ref;
}
