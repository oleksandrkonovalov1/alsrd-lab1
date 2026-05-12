# alsrd-lab1

Лабораторна робота №1 з дисципліни "Алгоритми та структури даних"

**Тема:** Вивчення часових характеристик алгоритмів

**Варіант 1:** Порівняння лінійного та двійкового пошуку

## Автор

- **Коновалов Олександр**, група ПЗПІ-25-6, oleksandr.konovalov1@nure.ua

## Технології

- TypeScript
- Node.js (tsx)
- docx (генерація звіту)
- VCS: Git + GitHub

## Опис проєкту

Порівняння часових характеристик трьох алгоритмів пошуку:
- Лінійний пошук — O(n)
- Двійковий пошук (ітеративний) — O(log n)
- Двійковий пошук (рекурсивний) — O(log n)

Програма виконує бенчмарк на масивах різного розміру (1 000 — 10 000 000 елементів), вимірює час виконання та порівнює з теоретичними оцінками.

## Запуск

```bash
git clone https://github.com/oleksandrkonovalov1/alsrd-lab1.git
cd alsrd-lab1
npm install

# Запустити бенчмарк
npm start

# Згенерувати звіт
npm run report
```

## Структура

```
alsrd-lab1/
├── src/
│   ├── algorithms/
│   │   ├── linear-search.ts
│   │   ├── binary-search-iterative.ts
│   │   └── binary-search-recursive.ts
│   ├── benchmark.ts
│   └── main.ts
├── reports/lr1/
│   ├── generate-report.mjs
│   ├── screenshots/
│   └── Звіт_ЛР1_Коновалов_ПЗПІ-25-6.docx
├── package.json
├── tsconfig.json
├── .gitignore
├── README.md
└── LICENSE
```

## Ліцензія

MIT License
