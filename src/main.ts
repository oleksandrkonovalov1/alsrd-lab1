import { runBenchmarks, formatResults } from "./benchmark.ts";

console.log("Лабораторна робота №1 — Часові характеристики алгоритмів пошуку");
console.log("Варіант 1: Лінійний пошук vs Двійковий пошук (ітеративний та рекурсивний)");
console.log();

console.log("Запуск бенчмарків...\n");
const results = runBenchmarks();
console.log(formatResults(results));
console.log();
console.log("Теоретичні оцінки:");
console.log("  Лінійний пошук:             O(n)");
console.log("  Двійковий пошук (ітератив): O(log n)");
console.log("  Двійковий пошук (рекурсія): O(log n)");
