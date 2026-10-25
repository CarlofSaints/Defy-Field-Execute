// Base columns found by header name, so an export that adds or moves a column
// still reads; the fallbacks are the positions Perigee and Field Goose use today.
export function baseCols(headers: readonly unknown[]) {
  const ix = (name: string, fallback: number) => {
    const i = headers.findIndex(h => String(h ?? '').trim().toLowerCase() === name);
    return i >= 0 ? i : fallback;
  };
  return {
    first: ix('first name', 2), last: ix('last name', 3), store: ix('store', 6),
    code: ix('store code', 7), province: ix('province', 8), date: ix('date', 9),
  };
}
