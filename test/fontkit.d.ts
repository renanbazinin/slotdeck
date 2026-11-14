// fontkit ships no type declarations; the tests and scripts only hand the module to the adapter.
declare module 'fontkit' {
  export function create(buffer: Uint8Array, postscriptName?: string): unknown;
}
