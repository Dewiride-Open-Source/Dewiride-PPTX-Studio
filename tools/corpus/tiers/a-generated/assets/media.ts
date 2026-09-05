/**
 * The two media files `a24-media` needs, authored here rather than found.
 *
 * The corpus rule is that a fixture is self-authored, CC0, public-domain
 * government or a synthetic probe - so a media deck cannot start by fetching a
 * sample clip, and on this machine it must not go looking for one either. Both
 * of these are written byte by byte from constants, are a few kilobytes, and
 * are a pure function of this file, which is what `C-REGEN` needs.
 *
 * ## The audio is real
 *
 * `probeWav()` is a quarter of a second of a 440 Hz sine at 8 kHz, mono, 8-bit
 * unsigned PCM - the simplest thing WAV can express and a file every decoder on
 * earth plays. Nothing about it is a stub.
 *
 * ## The video is a container and nothing else
 *
 * `probeMp4()` is a structurally valid ISO base media file - an `ftyp` box
 * declaring `isom`/`mp42`, and a `moov` box holding one `mvhd` - with **no
 * tracks and no coded samples**. It is a real MP4 in the sense that a parser
 * walks it to the end and finds nothing to play, which is exactly what it
 * claims to be.
 *
 * That is a deliberate limit and it is written down rather than glossed: coded
 * video means an encoder, an encoder means either a dependency or a download,
 * and neither is available here. What `a24-media` is the probe for is the
 * *markup* - `a:videoFile/@r:link` on an embedded part, the `p14:media` second
 * relationship to the same bytes, the poster frame, `p:stSnd`, and `p:audio`
 * and `p:video` inside a timeline. None of that needs a frame to decode. A
 * clip that actually plays is a Tier B deck, and `ROSTER.md` says so.
 */

const encoder = new TextEncoder();

function tag(name: string): Uint8Array {
  return encoder.encode(name);
}

function u32(value: number): Uint8Array {
  return new Uint8Array([
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff,
  ]);
}

function u16(value: number): Uint8Array {
  return new Uint8Array([(value >>> 8) & 0xff, value & 0xff]);
}

function u32le(value: number): Uint8Array {
  return new Uint8Array([
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  ]);
}

function u16le(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff]);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  let length = 0;
  for (const part of parts) length += part.length;
  const out = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

// ----------------------------------------------------------------- the WAV

const SAMPLE_RATE = 8000;
const TONE_HZ = 440;
const DURATION_SAMPLES = SAMPLE_RATE / 4;

/**
 * A quarter-second 440 Hz tone: RIFF, `fmt ` and `data`, and nothing else.
 *
 * 8-bit PCM samples are **unsigned**, centred on 128, which is the one thing
 * about WAV that catches people who have only written 16-bit.
 */
export function probeWav(): Uint8Array {
  const samples = new Uint8Array(DURATION_SAMPLES);
  for (let i = 0; i < DURATION_SAMPLES; i++) {
    const value = Math.sin((2 * Math.PI * TONE_HZ * i) / SAMPLE_RATE);
    samples[i] = Math.max(0, Math.min(255, Math.round(128 + value * 100)));
  }

  const fmt = concat([
    tag('fmt '),
    u32le(16),
    u16le(1), // PCM
    u16le(1), // mono
    u32le(SAMPLE_RATE),
    u32le(SAMPLE_RATE), // byte rate: rate * channels * bytesPerSample
    u16le(1), // block align
    u16le(8), // bits per sample
  ]);
  const data = concat([tag('data'), u32le(samples.length), samples]);
  const body = concat([tag('WAVE'), fmt, data]);
  return concat([tag('RIFF'), u32le(body.length), body]);
}

// ----------------------------------------------------------------- the MP4

/** An ISO-BMFF box: a big-endian length that includes the header, then the tag. */
function box(name: string, payload: Uint8Array): Uint8Array {
  return concat([u32(payload.length + 8), tag(name), payload]);
}

/** The 3x3 unity matrix ISO/IEC 14496-12 wants, in 16.16 and 2.30 fixed point. */
const UNITY_MATRIX = concat([
  u32(0x00010000),
  u32(0),
  u32(0),
  u32(0),
  u32(0x00010000),
  u32(0),
  u32(0),
  u32(0),
  u32(0x40000000),
]);

/** A valid MP4 container with no tracks. See the file comment. */
export function probeMp4(): Uint8Array {
  const ftyp = box('ftyp', concat([tag('isom'), u32(0x200), tag('isom'), tag('mp42')]));

  const mvhd = box(
    'mvhd',
    concat([
      u32(0), // version 0, no flags
      u32(0), // creation time - zero, so the bytes never move
      u32(0), // modification time
      u32(1000), // timescale: one millisecond
      u32(0), // duration: nothing to play
      u32(0x00010000), // rate 1.0
      u16(0x0100), // volume: 8.8 fixed point, so 1.0 is 0x0100
      u16(0), // reserved
      u32(0),
      u32(0), // reserved
      UNITY_MATRIX,
      concat([u32(0), u32(0), u32(0), u32(0), u32(0), u32(0)]), // pre_defined
      u32(2), // next track id
    ]),
  );

  return concat([ftyp, box('moov', mvhd)]);
}
