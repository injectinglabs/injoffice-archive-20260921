import type { NativeWasmClient, NativeWasmWorker } from '../src/index.js'

declare const workerFactory: () => NativeWasmWorker
declare const client: NativeWasmClient

void workerFactory
void client
