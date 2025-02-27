/* global console, window, clearTimeout */

/**
 * 防抖函数
 * @param func 需要防抖的函数
 * @param wait 等待时间（毫秒）
 */
export function debounce<T>(func: (...args: any[]) => Promise<T>, wait: number): (...args: any[]) => Promise<T> {
  let timeoutId: number | null = null;
  console.log(`创建防抖函数，等待时间: ${wait}ms`);

  return async (...args: any[]): Promise<T> => {
    console.log("防抖函数被调用");
    return new Promise((resolve) => {
      if (timeoutId !== null) {
        console.log("清除之前的定时器");
        clearTimeout(timeoutId);
      }

      console.log(`设置新的定时器，等待${wait}ms后执行`);
      timeoutId = window.setTimeout(async () => {
        console.log("定时器触发，执行原函数");
        try {
          const result = await func(...args);
          console.log("原函数执行成功");
          resolve(result);
        } catch (error) {
          console.error("原函数执行失败:", error);
          throw error;
        }
      }, wait);
    });
  };
}
