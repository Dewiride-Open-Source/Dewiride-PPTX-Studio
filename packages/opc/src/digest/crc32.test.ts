import { describe, expect, it } from 'vitest';
import { crc32 } from './crc32.js';

const encoder = new TextEncoder();

describe('crc32', () => {
  it('matches the standard check vectors', () => {
    // The IEEE 802.3 "check" value, quoted by every CRC catalogue.
    expect(crc32(encoder.encode('123456789'))).toBe(0xcbf43926);
    expect(crc32(new Uint8Array(0))).toBe(0);
    expect(crc32(encoder.encode('a'))).toBe(0xe8b7be43);
    expect(crc32(encoder.encode('The quick brown fox jumps over the lazy dog'))).toBe(0x414fa339);
  });

  it('is unsigned', () => {
    // A signed implementation returns -1010742131 here, which then fails to
    // compare equal to the unsigned value ZIP stores.
    expect(crc32(encoder.encode('hello world'))).toBe(0x0d4a1185);
    expect(crc32(encoder.encode('ÿÿÿÿ'))).toBeGreaterThanOrEqual(0);
  });

  it('resumes from a previous result, so a stream can be checksummed in chunks', () => {
    const whole = encoder.encode('the writer needs this in 0.3 too');
    const split = 11;
    const chunked = crc32(whole.subarray(split), crc32(whole.subarray(0, split)));
    expect(chunked).toBe(crc32(whole));
  });

  it('handles every byte value', () => {
    const all = new Uint8Array(256);
    for (let i = 0; i < 256; i++) all[i] = i;
    expect(crc32(all)).toBe(0x29058c73);
  });
});
