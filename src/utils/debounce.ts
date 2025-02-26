/**
 * 防抖函数
 * @param func 需要防抖的函数
 * @param wait 等待时间（毫秒）
 */
export function debounce<T>(
  func: (...args: any[]) => Promise<T>,
  wait: number
): (...args: any[]) => Promise<T> {
  let timeoutId: number | null = null;

  return async (...args: any[]): Promise<T> => {
    return new Promise((resolve) => {
      if (timeoutId !== null) {
        clearTimeout(timeoutId);
      }

      timeoutId = window.setTimeout(async () => {
        const result = await func(...args);
        resolve(result);
      }, wait);
    });
  };
}