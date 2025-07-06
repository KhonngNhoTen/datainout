import { Piscina } from "piscina";

export function createWorkerPool(maxThreads: number) {
  const workerPool = new Piscina({
    maxThreads,
    filename: "./Worker.js",
  });
  return workerPool;
}
