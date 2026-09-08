#!/usr/bin/env node
// Node >=22.18 supports the type-only TypeScript used by this local host.
import { createServer } from 'node:http'
import { handleAgentProposalRequest } from '../apps/playground/agentProposalHost.ts'

const port = Number(process.env.INJOFFICE_AGENT_HOST_PORT ?? 3102)
if (!Number.isInteger(port) || port < 1024 || port > 65535) throw new Error('Invalid INJOFFICE_AGENT_HOST_PORT')
const server = createServer((req, res) => {
  void handleAgentProposalRequest(req, res).then((handled) => {
    if (!handled) { res.statusCode = 404; res.end('Proposal API host; no UI is served here.\n') }
  }).catch(() => { res.statusCode = 500; res.end('Proposal host error\n') })
})
server.requestTimeout = 20_000
server.headersTimeout = 10_000
server.listen(port, '127.0.0.1', () => {
  console.log(`Local proposal API: http://127.0.0.1:${port}/api/agent/proposal-status`)
  console.log('No model is called until an explicitly consented POST. Configure the endpoint server-side.')
})
