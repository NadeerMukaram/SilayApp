# Breast Exam Checker

A React Native (Expo) Android app that imports clinical Excel sheets and checks the **Breast** column on the **last data row**.

## How it works

1. Import an `.xlsx` or `.xls` file (clinical exam format with many columns).
2. The app finds the **Breast** column header.
3. It reads the value from the **last row** with data.
4. Output:
   - **✓** if Breast = `UR`
   - **⚠** if Breast is anything else (or empty)

## Run on Android

```bash
npm install
npm run android
```

Or:

```bash
npx expo start
```

## Sample file

```bash
npm run sample
```

Creates `sample-data/sample-clinical.xlsx` based on your clinical sheet format.

## Setup

Requires Expo Go on your Android device, or an Android emulator with the dev server running.
