import { Schema } from 'prosemirror-model'
import { nodes as basicNodes } from 'prosemirror-schema-basic'
import { Step } from 'prosemirror-transform'

/** Tiny doc/paragraph/text schema. Marks and extra blocks stay out of this proof. */
export const docsCollabSchema = new Schema({
  nodes: {
    doc: basicNodes.doc,
    paragraph: basicNodes.paragraph,
    text: basicNodes.text,
  },
})

export function seedDocsCollabDoc() {
  return docsCollabSchema.node('doc', null, [
    docsCollabSchema.node('paragraph', null, [
      docsCollabSchema.text('Shared document. Type here, then watch the same characters appear in the other tab.'),
    ]),
    docsCollabSchema.node('paragraph'),
  ])
}

export function stepsFromOps(ops: unknown[]): Step[] {
  return ops.map((op) => Step.fromJSON(docsCollabSchema, op))
}

export function opsFromSteps(steps: readonly Step[]): unknown[] {
  return steps.map((step) => step.toJSON())
}
