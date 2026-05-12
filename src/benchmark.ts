import { performance } from "node:perf_hooks";
import { linearSearch } from "./algorithms/linear-search.ts";
import { binarySearchIterative } from "./algorithms/binary-search-iterative.ts";
import { binarySearchRecursive } from "./algorithms/binary-search-recursive.ts";

type SearchFn = (arr: number[], target: number) => number;

interface BenchmarkResult {
  n: number;
  linearMs: number;
  binaryIterMs: number;
  binaryRecMs: number;
}

function measureAverage(fn: SearchFn, arr: number[], target: number, iterations: number): number {
  for (let i = 0; i < Math.min(iterations, 10); i++) {
    fn(arr, target);
  }

  const start = performance.now();
  for (let i = 0; i < iterations; i++) {
    fn(arr, target);
  }
  const elapsed = performance.now() - start;
  return elapsed / iterations;
}

function chooseIterations(n: number, isFast: boolean): number {
  if (isFast) {
    if (n <= 10_000) return 100_000;
    if (n <= 1_000_000) return 50_000;
    return 10_000;
  }
  if (n <= 10_000) return 10_000;
  if (n <= 100_000) return 1_000;
  if (n <= 1_000_000) return 100;
  return 10;
}

export function runBenchmarks(): BenchmarkResult[] {
  const sizes = [1_000, 10_000, 100_000, 1_000_000, 10_000_000];
  const results: BenchmarkResult[] = [];

  for (const n of sizes) {
    const arr = Array.from({ length: n }, (_, i) => i);
    const target = n;

    const linearIter = chooseIterations(n, false);
    const binaryIter = chooseIterations(n, true);

    const linearMs = measureAverage(linearSearch, arr, target, linearIter);
    const binaryIterMs = measureAverage(binarySearchIterative, arr, target, binaryIter);
    const binaryRecMs = measureAverage(binarySearchRecursive, arr, target, binaryIter);

    results.push({ n, linearMs, binaryIterMs, binaryRecMs });
  }

  return results;
}

export function formatResults(results: BenchmarkResult[]): string {
  const header = "n".padStart(12) + "Linear (ms)".padStart(16) + "Binary iter (ms)".padStart(20) + "Binary rec (ms)".padStart(20);
  const separator = "─".repeat(header.length);

  const rows = results.map((r) => {
    const nStr = r.n.toLocaleString("uk-UA").padStart(12);
    const lin = r.linearMs.toFixed(4).padStart(16);
    const binI = r.binaryIterMs.toFixed(4).padStart(20);
    const binR = r.binaryRecMs.toFixed(4).padStart(20);
    return `${nStr}${lin}${binI}${binR}`;
  });

  return [header, separator, ...rows].join("\n");
}
