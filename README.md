# ExcelJS React Hook Example

## ℹ️ Overview

An example React project demonstrating how to generate Excel spreadsheets using **ExcelJS** via a **custom, reusable React hook**.

This project focuses on:

- Encapsulating ExcelJS logic inside a hook
- Providing a clean, typed API for consumers
- Supporting flexible column and worksheet configuration
- Keeping UI code simple and declarative

## 💡 What This Project Demonstrates

- A generic `useExcelDownload<T>()` hook
- Strong TypeScript typing for rows, columns, and workbooks
- Optional column configuration (labels, widths, renderers)
- Multiple worksheet support
- Client-side Excel file generation (no backend required)

## 🚀 Running the Project

```bash
npm install
npm run dev
```

Open the app in your browser and trigger a download to generate an `.xlsx` file.

## 📄 Core Concept

Instead of calling ExcelJS APIs directly from components, this project wraps that logic inside a **custom hook**.

This approach:

- Keeps components focused on UI
- Improves reuse across projects
- Makes spreadsheet generation easier to test and evolve

## 🛠 Example Usage

### 1. Define Column Configuration (Optional)

```ts
const columns: Array<ExcelDownloadColumn<MyModel>> = [
  {
    key: 'id',
  },
  {
    key: 'firstName',
    label: 'First Name',
    width: 20,
  },
  {
    key: 'lastName',
    label: 'Last Name',
    width: 'auto',
    render: (lastName: string) => lastName.toUpperCase(),
  },
  {
    key: 'address',
    altKey: 'street1',
    label: 'Address Line 1',
    render: (address: Address) => address.street1,
  },
];
```

### 2. Initialize the Custom Hook

```ts
const [downloadExcel] = useExcelDownload<MyModel>();
```

### 3. Build the Workbook Configuration

```ts
const workbook: ExcelDownloadWorkbook<MyModel> = {
  filename: 'my-excel-file',
  worksheets: [
    {
      list: myDataToConvert,
      worksheetName: 'My Worksheet',
      columns,
      defaultColumnWidth: 'auto',
    },
  ],
};
```

> ℹ️ Worksheet names must follow Excel naming rules  
> https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9

### 4. Trigger the Download

```ts
const handleDownload = useCallback(() => {
  downloadExcel(workbook);
}, [downloadExcel]);
```

```tsx
<button onClick={handleDownload}>Download</button>
```

## 🎯 Design Goals

- Hide ExcelJS implementation details
- Prefer configuration over imperative code
- Make column behavior explicit and readable
- Support common spreadsheet use cases without over-engineering

## 📌 Notes

- This is a **client-side** solution intended for moderate data sizes
- Large datasets may require streaming or server-side generation
- ExcelJS is used directly—no wrappers or abstractions beyond the hook

## 📄 License

MIT
