package pptxpatch

// NativeEvaluatedGeometry is source-evaluated DrawingML, never a browser-side
// formula program. Coordinates are rounded once to local integer EMU.
type NativeEvaluatedGeometry struct {
	Profile  string                 `json:"profile"`
	TextRect NativeGeometryTextRect `json:"textRect"`
	Paths    []NativeGeometryPath   `json:"paths"`
}
type NativeGeometryTextRect struct {
	X  int64 `json:"x"`
	Y  int64 `json:"y"`
	CX int64 `json:"cx"`
	CY int64 `json:"cy"`
}
type NativeGeometryPath struct {
	FillMode string                  `json:"fillMode"`
	Stroke   bool                    `json:"stroke"`
	Commands []NativeGeometryCommand `json:"commands"`
}

// Command fields are pointers so exact zero is retained while absent fields
// remain absent. Contract validation will reject fields irrelevant to Kind.
type NativeGeometryCommand struct {
	Kind      string `json:"kind"`
	X         *int64 `json:"x,omitempty"`
	Y         *int64 `json:"y,omitempty"`
	X1        *int64 `json:"x1,omitempty"`
	Y1        *int64 `json:"y1,omitempty"`
	X2        *int64 `json:"x2,omitempty"`
	Y2        *int64 `json:"y2,omitempty"`
	RX        *int64 `json:"rx,omitempty"`
	RY        *int64 `json:"ry,omitempty"`
	LargeArc  *bool  `json:"largeArc,omitempty"`
	Clockwise *bool  `json:"clockwise,omitempty"`
}
