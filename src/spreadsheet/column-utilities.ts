export function a1Col(colIndex: number): string {
  let col = '';
  let i = colIndex;
  while (i > 0) {
    const rem = (i - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    i = Math.floor((i - 1) / 26);
  }
  return col;
}
