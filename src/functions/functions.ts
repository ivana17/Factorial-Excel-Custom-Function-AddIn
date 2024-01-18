/* global clearInterval, console, CustomFunctions, setInterval */

declare global {
  interface Window {
    direction: string;
  }
}

window.direction = "Row";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Facts a number.
 * @customfunction FACTORIALROW
 * @param num A number.
 * @returns The factorial of the number.
 */
export function fact(num: number): number[][] {
  const contents: number[][] = [];
  let row: number[] = [];

  if (window.direction === "Row") contents.push(row);

  for (let i = 1, fact = 1; i <= num; i++) {
    fact *= i;

    if (window.direction === "Row") {
      row.push(fact);
    } else if (window.direction === "Column") {
      contents.push([fact]);
    }
  }

  return contents;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
