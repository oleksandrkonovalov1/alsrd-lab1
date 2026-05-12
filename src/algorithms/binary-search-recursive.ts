function binarySearchHelper(
  arr: number[],
  target: number,
  low: number,
  high: number,
): number {
  if (low > high) {
    return -1;
  }

  const mid = Math.floor((low + high) / 2);
  if (arr[mid] === target) {
    return mid;
  }
  if (arr[mid] < target) {
    return binarySearchHelper(arr, target, mid + 1, high);
  }
  return binarySearchHelper(arr, target, low, mid - 1);
}

export function binarySearchRecursive(arr: number[], target: number): number {
  return binarySearchHelper(arr, target, 0, arr.length - 1);
}
