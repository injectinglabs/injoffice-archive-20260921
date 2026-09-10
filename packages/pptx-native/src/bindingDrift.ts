import type {
  NativeAnimation,
  NativeAsset,
  NativeChartElement,
  NativeCompatibility,
  NativeConnectorElement,
  NativeDiagnostic,
  NativeDiagnosticScope,
  NativeGroupElement,
  NativeOpaqueChart,
  NativeParagraph,
  NativePassthroughRef,
  NativePictureElement,
  NativePictureCrop,
  NativePptxDeck,
  NativeShapeElement,
  NativeSize,
  NativeSlide,
  NativeSourceAnchor,
  NativeStroke,
  NativeTable,
  NativeTableBorder,
  NativeTableCell,
  NativeTableElement,
  NativeTextElement,
  NativeTextBodyLayout,
  NativeTextRun,
  NativeTransition,
  NativeTransform,
} from './types'
import { PPTX_NATIVE_OBJECT_BINDINGS } from './schema.generated'

type Equal<Left, Right> =
  (<Value>() => Value extends Left ? 1 : 2) extends
  (<Value>() => Value extends Right ? 1 : 2) ? true : false
type Assert<Value extends true> = Value
type RequiredKeys<Value> = {
  [Key in keyof Value]-?: object extends Pick<Value, Key> ? never : Key
}[keyof Value]
type BindingName = keyof typeof PPTX_NATIVE_OBJECT_BINDINGS
type SchemaFields<Name extends BindingName> = (typeof PPTX_NATIVE_OBJECT_BINDINGS)[Name]['properties'][number]
type SchemaRequired<Name extends BindingName> = (typeof PPTX_NATIVE_OBJECT_BINDINGS)[Name]['required'][number]
type Fields<Name extends BindingName, Value> = Equal<keyof Value, SchemaFields<Name>>
type Required<Name extends BindingName, Value> = Equal<RequiredKeys<Value>, SchemaRequired<Name>>

// These declarations intentionally emit no JavaScript. Typechecking fails if
// a handwritten interface drifts from the canonical schema's generated field
// or required-property manifest.
type NativePptxBindingChecks = [
  Assert<Fields<'NativePptxDeck', NativePptxDeck>>, Assert<Required<'NativePptxDeck', NativePptxDeck>>,
  Assert<Fields<'NativeSize', NativeSize>>, Assert<Required<'NativeSize', NativeSize>>,
  Assert<Fields<'NativeTransform', NativeTransform>>, Assert<Required<'NativeTransform', NativeTransform>>,
  Assert<Fields<'NativeSourceAnchor', NativeSourceAnchor>>, Assert<Required<'NativeSourceAnchor', NativeSourceAnchor>>,
  Assert<Fields<'NativePassthroughRef', NativePassthroughRef>>, Assert<Required<'NativePassthroughRef', NativePassthroughRef>>,
  Assert<Fields<'NativeDiagnosticScope', NativeDiagnosticScope>>, Assert<Required<'NativeDiagnosticScope', NativeDiagnosticScope>>,
  Assert<Fields<'NativeDiagnostic', NativeDiagnostic>>, Assert<Required<'NativeDiagnostic', NativeDiagnostic>>,
  Assert<Fields<'NativeCompatibility', NativeCompatibility>>, Assert<Required<'NativeCompatibility', NativeCompatibility>>,
  Assert<Fields<'NativeAsset', NativeAsset>>, Assert<Required<'NativeAsset', NativeAsset>>,
  Assert<Fields<'NativeTextRun', NativeTextRun>>, Assert<Required<'NativeTextRun', NativeTextRun>>,
  Assert<Fields<'NativeParagraph', NativeParagraph>>, Assert<Required<'NativeParagraph', NativeParagraph>>,
  Assert<Fields<'NativeTextBodyLayout', NativeTextBodyLayout>>, Assert<Required<'NativeTextBodyLayout', NativeTextBodyLayout>>,
  Assert<Fields<'NativeStroke', NativeStroke>>, Assert<Required<'NativeStroke', NativeStroke>>,
  Assert<Fields<'NativeAnimation', NativeAnimation>>, Assert<Required<'NativeAnimation', NativeAnimation>>,
  Assert<Fields<'NativeTransition', NativeTransition>>, Assert<Required<'NativeTransition', NativeTransition>>,
  Assert<Fields<'NativeTableBorder', NativeTableBorder>>, Assert<Required<'NativeTableBorder', NativeTableBorder>>,
  Assert<Fields<'NativeTableCell', NativeTableCell>>, Assert<Required<'NativeTableCell', NativeTableCell>>,
  Assert<Fields<'NativeTable', NativeTable>>, Assert<Required<'NativeTable', NativeTable>>,
  Assert<Fields<'NativeOpaqueChart', NativeOpaqueChart>>, Assert<Required<'NativeOpaqueChart', NativeOpaqueChart>>,
  Assert<Fields<'NativeTextElement', NativeTextElement>>, Assert<Required<'NativeTextElement', NativeTextElement>>,
  Assert<Fields<'NativeShapeElement', NativeShapeElement>>, Assert<Required<'NativeShapeElement', NativeShapeElement>>,
  Assert<Fields<'NativeConnectorElement', NativeConnectorElement>>, Assert<Required<'NativeConnectorElement', NativeConnectorElement>>,
  Assert<Fields<'NativePictureElement', NativePictureElement>>, Assert<Required<'NativePictureElement', NativePictureElement>>,
  Assert<Fields<'NativePictureCrop', NativePictureCrop>>, Assert<Required<'NativePictureCrop', NativePictureCrop>>,
  Assert<Fields<'NativeTableElement', NativeTableElement>>, Assert<Required<'NativeTableElement', NativeTableElement>>,
  Assert<Fields<'NativeChartElement', NativeChartElement>>, Assert<Required<'NativeChartElement', NativeChartElement>>,
  Assert<Fields<'NativeGroupElement', NativeGroupElement>>, Assert<Required<'NativeGroupElement', NativeGroupElement>>,
  Assert<Fields<'NativeSlide', NativeSlide>>, Assert<Required<'NativeSlide', NativeSlide>>,
]

export type { NativePptxBindingChecks }
