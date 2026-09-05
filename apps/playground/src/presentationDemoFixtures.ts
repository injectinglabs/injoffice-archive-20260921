import type { DeckSpec } from '@injoffice/slides/authoring'

export const PRESENTATION_DEMO_TITLE = 'Northstar launch review'

export const PRESENTATION_DEMO_OUTLINE = `# Northstar launch review
Q3 operating brief · September 2026
Notes: Open with the customer outcome, then move to the operating evidence.
## The quarter at a glance
- **128%** of activation target
- Time-to-value improved from 11 to **7 days**
- Enterprise pilot conversion reached **64%**
- Support backlog fell **31%**
## What changed
- Guided onboarding now adapts to team size
- Admin setup is one workflow instead of four
- Usage alerts reach owners before adoption stalls
## Next 30 days
- Ship audit-log export to every enterprise pilot
- Close accessibility findings from the release review
- Publish migration playbooks for the top three source systems
> Make the safe path the fast path. — Launch review principle
# Decision
Approve the October rollout with weekly adoption and reliability checkpoints.
`

export function makeAuthoredPresentationDemo(
  title: string,
  subtitle: string,
  theme: string,
): DeckSpec {
  return {
    id: 'northstar-launch-review',
    title,
    theme,
    slides: [
      {
        id: 'opening-slide',
        kind: 'title',
        eyebrow: 'Operating review · Q3 2026',
        title,
        subtitle,
        transition: { kind: 'fade' },
      },
      {
        id: 'evidence-slide',
        kind: 'bullets',
        eyebrow: 'Adoption evidence',
        title: 'The launch is converting into durable usage',
        bullets: [
          '**128%** of activation target',
          'Time-to-value improved from 11 to **7 days**',
          'Enterprise pilot conversion reached **64%**',
          'Support backlog fell **31%**',
        ],
        transition: { kind: 'push', direction: 'left' },
      },
      {
        id: 'plan-slide',
        kind: 'two-col',
        eyebrow: 'October plan',
        title: 'Scale adoption without trading away trust',
        colTitles: ['Product', 'Operations'],
        bullets: [
          'Audit-log export',
          'Adaptive onboarding',
          'Migration playbooks',
        ],
        bulletsRight: [
          'Weekly reliability review',
          'Accessibility closeout',
          'Named pilot owners',
        ],
        transition: { kind: 'wipe', direction: 'left' },
      },
      {
        id: 'decision-slide',
        kind: 'closing',
        title: 'Approve the October rollout',
        subtitle: 'Proceed with weekly adoption and reliability checkpoints.',
        transition: { kind: 'fade' },
      },
    ],
  }
}
