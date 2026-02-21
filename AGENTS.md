# Agent Information

## How to Start the Web App (for Development)

If you need to verify operations on the Web UI or test editor features, start the development server using the following commands:

```bash
# Install dependencies (only required the first time)
npm install

# Start the local development server (default: http://localhost:5173/)
npm run dev
```

## How to Run Automated Tests

This project includes a TypeScript-based unit testing environment for verifying the logic of VBA code.
When an agent makes changes to the code, please run the following command to ensure there are no regressions:

```bash
# Bundle and run tests (AST construction and verification)
npx esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
```
