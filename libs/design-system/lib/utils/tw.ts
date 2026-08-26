export function tw(
  template: TemplateStringsArray,
): string;
export function tw(className: string): string;
/**
 *
 * @param val
 */
export function tw(
  val: string | TemplateStringsArray,
) {
  if (typeof val === 'string') {
    return val;
  }

  return String.raw({ raw: val });
}
