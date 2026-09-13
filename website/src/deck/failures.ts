import type { Failure } from '@/design/state';

/** What a visitor can do about the errors a dropped deck most often raises. */
const ADVICE: Readonly<Record<string, string>> = {
  ERR_NOT_A_ZIP:
    'This is not a PowerPoint file. A .pptx is a ZIP archive; a .ppt from PowerPoint 2003 is not, and PPTX Studio does not read it.',
  ERR_ENCRYPTED_PACKAGE:
    'The file is password-protected. PPTX Studio cannot open encrypted decks; remove the password in PowerPoint and try again.',
  ERR_TRUNCATED:
    'The file ends before its central directory. It was probably cut off in transfer; download it again.',
  ERR_ARCHIVE_TOO_LARGE:
    'The archive is over the reader’s size limit. The limits exist so that a hostile file cannot take the tab down.',
  ERR_RATIO_EXCEEDED:
    'An entry inflates far beyond its compressed size - the shape of a zip bomb - so the reader stopped.',
  MODEL_NO_PRESENTATION:
    'The package has no presentation part. It may be a Word or Excel file with the wrong extension.',
  TEXT_NO_CANVAS:
    'Text measurement needs a canvas, which this browser did not provide. Text layout is unavailable here.',
  TEXT_AUTONUMBER_UNMEASURED:
    'A paragraph uses one of the two autonumber schemes PowerPoint shapes into glyph ids, which the text engine refuses rather than guesses.',
};

/** Read the code a typed error carries, and say in plain words what it means. */
export function describeFailure(cause: unknown): Failure {
  if (cause instanceof Error) {
    const code = 'code' in cause && typeof cause.code === 'string' ? cause.code : undefined;
    return { code, message: cause.message, advice: code === undefined ? undefined : ADVICE[code] };
  }
  return { message: String(cause) };
}
