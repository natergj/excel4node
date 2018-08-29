import { PassThrough } from 'stream';

export function getDataStream() {
  return new PassThrough();
}
