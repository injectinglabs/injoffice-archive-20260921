/* Copyright 2026 Injecting Inc. SPDX-License-Identifier: Apache-2.0 */
// The profile is fixed before loading the shared worker transport. Incoming
// messages cannot enable extraction or mutation on this worker.
Object.defineProperty(self, 'xlsxSourceStylePreview', { value: true })
// The package build appends xlsxnative.worker.js here, keeping this worker
// self-contained when bundlers fingerprint its asset URL.
