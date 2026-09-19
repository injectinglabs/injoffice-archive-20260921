import { calculateLocalWorkbook, type LocalCalculationInput } from '../../../packages/formulas/src/localWorkbookCalculation';
self.onmessage = async (event: MessageEvent<LocalCalculationInput>) => {
  try { self.postMessage({ result: await calculateLocalWorkbook(event.data) }); }
  catch (reason) { self.postMessage({ error: reason instanceof Error ? reason.message : String(reason) }); }
};
