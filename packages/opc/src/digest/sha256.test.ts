import { describe, expect, it } from 'vitest';
import { sha256, sha256Hex, sha256HexOfText } from './sha256.js';

/**
 * A hash nobody checked against something else is a hash that agrees with
 * itself.
 *
 * The first block below is FIPS 180-4's own appendix: the empty message, the
 * one-block `abc`, the 56-character example that no longer fits in one block,
 * and the million-`a` long-message vector. Nothing here was produced by this
 * implementation.
 *
 * The second block pins the four lengths where the padding changes shape, and
 * those digests come from Node's `crypto.createHash('sha256')` - OpenSSL, a
 * genuinely independent implementation - rather than from the standard, which
 * does not publish a vector at every boundary.
 */

const ascii = (text: string): Uint8Array => new TextEncoder().encode(text);

describe('the FIPS 180-4 vectors', () => {
  it('digests the empty message', () => {
    expect(sha256Hex(new Uint8Array(0))).toBe(
      'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    );
  });

  it('digests "abc" - the one-block example', () => {
    expect(sha256Hex(ascii('abc'))).toBe(
      'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    );
  });

  it('digests the 56-character example - two blocks, because the length no longer fits', () => {
    const message = 'abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq';
    expect(message).toHaveLength(56);
    expect(sha256Hex(ascii(message))).toBe(
      '248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1',
    );
  });

  it('digests a million "a" - the long-message vector', () => {
    expect(sha256Hex(new Uint8Array(1_000_000).fill(0x61))).toBe(
      'cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0',
    );
  });
});

describe('the lengths where the padding changes shape', () => {
  /**
   * 55 is the last length whose `0x80` marker and eight-byte length still fit
   * in the same block; 56 is the first that needs a second block; 64 is a whole
   * number of blocks with nothing left over, so the tail is padding alone; 119
   * and 120 are the same boundary one block further along, which is where an
   * implementation that pads relative to the wrong offset first disagrees.
   */
  const vectors: readonly (readonly [number, string])[] = [
    [55, '9f4390f8d30c2dd92ec9f095b65e2b9ae9b0a925a5258e241c9f1e910f734318'],
    [56, 'b35439a4ac6f0948b6d6f9e3c6af0f5f590ce20f1bde7090ef7970686ec6738a'],
    [63, '7d3e74a05d7db15bce4ad9ec0658ea98e3f06eeecf16b4c6fff2da457ddc2f34'],
    [64, 'ffe054fe7ae0cb6dc65c3af9b61d5209f439851db43d0ba5997337df154668eb'],
    [65, '635361c48bb9eab14198e76ea8ab7f1a41685d6ad62aa9146d301d4f17eb0ae0'],
    [119, '31eba51c313a5c08226adf18d4a359cfdfd8d2e816b13f4af952f7ea6584dcfb'],
    [120, '2f3d335432c70b580af0e8e1b3674a7c020d683aa5f73aaaedfdc55af904c21c'],
  ];

  for (const [length, expected] of vectors) {
    it('digests ' + String(length) + ' bytes of "a"', () => {
      expect(sha256Hex(new Uint8Array(length).fill(0x61))).toBe(expected);
    });
  }
});

describe('the shape of the output', () => {
  it('returns thirty-two bytes', () => {
    expect(sha256(ascii('anything'))).toHaveLength(32);
  });

  it('spells the hex form in sixty-four lowercase digits', () => {
    expect(sha256Hex(ascii('anything'))).toMatch(/^[0-9a-f]{64}$/);
  });

  it('does not read past a subarray view', () => {
    // A part read out of an archive is a view onto the archive buffer, so a
    // digest that used the backing store's length rather than the view's would
    // silently hash the neighbouring entry as well.
    const backing = ascii('__abc__');
    expect(sha256Hex(backing.subarray(2, 5))).toBe(sha256Hex(ascii('abc')));
  });

  it('encodes text as UTF-8 before digesting', () => {
    // Not the UTF-16 code units the string is stored in: two bytes for the
    // pound sign, one each for the ASCII.
    expect(sha256HexOfText('£')).toBe(sha256Hex(new Uint8Array([0xc2, 0xa3])));
    expect(sha256HexOfText('abc')).toBe(sha256Hex(ascii('abc')));
  });

  it('separates messages that differ in one bit', () => {
    expect(sha256Hex(new Uint8Array([0x00]))).not.toBe(sha256Hex(new Uint8Array([0x01])));
  });
});
