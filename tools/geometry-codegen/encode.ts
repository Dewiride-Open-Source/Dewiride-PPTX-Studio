import type { PresetCommand, PresetShape } from '../../packages/geometry/src/types.ts';

/**
 * The write half of the encoding `packages/geometry/src/decode.ts` reads.
 *
 * The grammar is documented there, once, next to the code that has to survive
 * in a browser. What lives here is the inverse and the safety check: every
 * token is asserted to contain none of the five separators before it is
 * written, so a future POI release that introduces a comma into an operand
 * stops the build rather than quietly splitting one shape into two.
 */

/** The characters the encoding reserves. None occurs in the source data today. */
const RESERVED = /[;|,!]|\s/;

function safe(token: string, what: string): string {
  if (token === '') throw new Error(`${what}: empty token`);
  if (RESERVED.test(token)) {
    throw new Error(`${what}: "${token}" contains a character the encoding reserves`);
  }
  return token;
}

/** `null` is written as a lone `-`, which no real operand can be. */
function maybe(value: string | null, what: string): string {
  return value === null ? '-' : safe(value, what);
}

function encodeCommand(command: PresetCommand): string {
  switch (command.kind) {
    case 'moveTo':
      return `M ${safe(command.to.x, 'moveTo')} ${safe(command.to.y, 'moveTo')}`;
    case 'lnTo':
      return `L ${safe(command.to.x, 'lnTo')} ${safe(command.to.y, 'lnTo')}`;
    case 'quadBezTo':
      return [
        'Q',
        safe(command.c1.x, 'quadBezTo'),
        safe(command.c1.y, 'quadBezTo'),
        safe(command.to.x, 'quadBezTo'),
        safe(command.to.y, 'quadBezTo'),
      ].join(' ');
    case 'cubicBezTo':
      return [
        'C',
        safe(command.c1.x, 'cubicBezTo'),
        safe(command.c1.y, 'cubicBezTo'),
        safe(command.c2.x, 'cubicBezTo'),
        safe(command.c2.y, 'cubicBezTo'),
        safe(command.to.x, 'cubicBezTo'),
        safe(command.to.y, 'cubicBezTo'),
      ].join(' ');
    case 'arcTo':
      return [
        'A',
        safe(command.wR, 'arcTo'),
        safe(command.hR, 'arcTo'),
        safe(command.stAng, 'arcTo'),
        safe(command.swAng, 'arcTo'),
      ].join(' ');
    case 'close':
      return 'Z';
  }
}

export function encodeShape(shape: PresetShape): string {
  const where = shape.name;

  const guides = (list: PresetShape['avLst']): string =>
    list
      .map((gd) => [safe(gd.name, where), ...gd.fmla.map((t) => safe(t, where))].join(' '))
      .join(',');

  const handles = shape.ahLst
    .map((handle) =>
      handle.kind === 'xy'
        ? [
            'X',
            maybe(handle.gdRefX, where),
            maybe(handle.minX, where),
            maybe(handle.maxX, where),
            maybe(handle.gdRefY, where),
            maybe(handle.minY, where),
            maybe(handle.maxY, where),
            safe(handle.pos.x, where),
            safe(handle.pos.y, where),
          ].join(' ')
        : [
            'P',
            maybe(handle.gdRefR, where),
            maybe(handle.minR, where),
            maybe(handle.maxR, where),
            maybe(handle.gdRefAng, where),
            maybe(handle.minAng, where),
            maybe(handle.maxAng, where),
            safe(handle.pos.x, where),
            safe(handle.pos.y, where),
          ].join(' '),
    )
    .join(',');

  const sites = shape.cxnLst
    .map((site) =>
      [safe(site.ang, where), safe(site.pos.x, where), safe(site.pos.y, where)].join(' '),
    )
    .join(',');

  const rect =
    shape.rect === null
      ? ''
      : [
          safe(shape.rect.l, where),
          safe(shape.rect.t, where),
          safe(shape.rect.r, where),
          safe(shape.rect.b, where),
        ].join(' ');

  const paths = shape.pathLst
    .map((path) => {
      const header = [
        path.w === 0 ? '-' : String(path.w),
        path.h === 0 ? '-' : String(path.h),
        path.fill === 'norm' ? '-' : safe(path.fill, where),
        path.stroke ? '-' : '0',
        path.extrusionOk ? '-' : '0',
      ].join(' ');
      return [header, ...path.commands.map(encodeCommand)].join(',');
    })
    .join('!');

  return [
    safe(shape.name, where),
    guides(shape.avLst),
    guides(shape.gdLst),
    handles,
    sites,
    rect,
    paths,
  ].join('|');
}

export function encodeBucket(shapes: readonly PresetShape[]): string {
  return shapes.map(encodeShape).join(';');
}
