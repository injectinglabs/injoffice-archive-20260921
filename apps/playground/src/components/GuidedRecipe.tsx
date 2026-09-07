import { useEffect, useState } from 'react'
import type { DemoDefinition } from '../demoRegistry'
import type { DemoRecipe } from '../demoRecipes'
import { surfaceHref } from '../route'
import './GuidedRecipe.css'

type RecipeAccent = DemoDefinition['accent']

export type GuidedRecipeProps = {
  recipe: DemoRecipe
  accent?: RecipeAccent
  className?: string
}

export function GuidedRecipe({ recipe, accent = 'blue', className = '' }: GuidedRecipeProps) {
  const [activeIndex, setActiveIndex] = useState(0)
  const [completed, setCompleted] = useState<ReadonlySet<string>>(() => new Set())
  const activeStep = recipe.steps[activeIndex] ?? recipe.steps[0]
  const completeCount = completed.size
  const allComplete = completeCount === recipe.steps.length

  useEffect(() => {
    setActiveIndex(0)
    setCompleted(new Set())
  }, [recipe.id])

  const completeActiveStep = () => {
    if (!activeStep) return
    setCompleted((current) => new Set(current).add(activeStep.id))
    setActiveIndex((current) => Math.min(current + 1, recipe.steps.length - 1))
  }

  const reset = () => {
    setActiveIndex(0)
    setCompleted(new Set())
  }

  return (
    <aside className={`guided-recipe guided-recipe--${accent}${className ? ` ${className}` : ''}`} aria-labelledby={`${recipe.id}-title`}>
      <header className="guided-recipe__header">
        <div>
          <p className="guided-recipe__kind">Try this · about {recipe.minutes} min</p>
          <h2 id={`${recipe.id}-title`}>{recipe.title}</h2>
          <p>{recipe.outcome}</p>
        </div>
        <div
          className="guided-recipe__progress"
          role="progressbar"
          aria-label="Recipe progress"
          aria-valuemin={0}
          aria-valuemax={recipe.steps.length}
          aria-valuenow={completeCount}
        >
          <strong>{completeCount}/{recipe.steps.length}</strong>
          <span>steps checked</span>
        </div>
      </header>

      <ol className="guided-recipe__rail" aria-label="Recipe steps">
        {recipe.steps.map((step, index) => {
          const isComplete = completed.has(step.id)
          const isActive = index === activeIndex
          return (
            <li key={step.id} data-complete={isComplete || undefined} data-active={isActive || undefined}>
              <button
                type="button"
                aria-current={isActive ? 'step' : undefined}
                aria-label={`${isComplete ? 'Checked step' : 'Step'} ${index + 1}: ${step.title}`}
                onClick={() => setActiveIndex(index)}
              >
                <span aria-hidden="true">{isComplete ? '✓' : index + 1}</span>
                <strong>{step.title}</strong>
              </button>
            </li>
          )
        })}
      </ol>

      {activeStep ? (
        <section className="guided-recipe__step" aria-live="polite" aria-labelledby={`${recipe.id}-${activeStep.id}-title`}>
          <div className="guided-recipe__instruction">
            <span>Step {activeIndex + 1}</span>
            <h3 id={`${recipe.id}-${activeStep.id}-title`}>{activeStep.title}</h3>
            <p>{activeStep.instruction}</p>
          </div>
          <div className="guided-recipe__evidence">
            <strong>Look for</strong>
            <p>{activeStep.evidence}</p>
          </div>
          <div className="guided-recipe__actions">
            <button className="guided-recipe__complete" type="button" onClick={completeActiveStep} disabled={completed.has(activeStep.id)}>
              {completed.has(activeStep.id) ? 'Step checked' : activeIndex === recipe.steps.length - 1 ? 'Finish checklist' : 'Mark done & continue'}
            </button>
            {completeCount > 0 ? <button className="guided-recipe__reset" type="button" onClick={reset}>Start over</button> : null}
            {allComplete ? <span role="status">Checklist complete. Check the demo’s actual results before relying on the output.</span> : null}
          </div>
        </section>
      ) : null}

      <footer className="guided-recipe__footer">
        <div>
          <strong>Source</strong>
          <ul>
            {recipe.sources.map((item) => (
              <li key={item.path}><a href={item.href} target="_blank" rel="noreferrer" title={item.path}>{item.label}<span aria-hidden="true"> ↗</span></a></li>
            ))}
          </ul>
        </div>
        <div>
          <strong>Continue exploring</strong>
          <ul>
            {recipe.related.map((item) => (
              <li key={item.surface}><a href={surfaceHref(item.surface)} title={item.reason}>{item.surface.replaceAll('-', ' ')}</a></li>
            ))}
          </ul>
        </div>
      </footer>
    </aside>
  )
}

export default GuidedRecipe
