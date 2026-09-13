package pptxpatch

import (
	"crypto/sha256"
	"fmt"
	"math/big"
	"strconv"
	"strings"
)

// A guide has at most 4096 numerator/denominator bits. Every arithmetic step
// combines at most three already bounded operands, limiting temporary growth.
const nativeGeometryMaxRationalBits = 4096

func (g nativeGeometryGuides) exactOperand(token string) *big.Rat {
	if _, exists := g.values[token]; exists {
		return g.exact[token]
	}
	value, err := strconv.ParseInt(token, 10, 64)
	if err != nil || !nativeGeometryFinite(float64(value)) {
		return nil
	}
	return big.NewRat(value, 1)
}
func nativeGeometryExactSqrt(value *big.Rat) *big.Rat {
	if value == nil || value.Sign() < 0 {
		return nil
	}
	n := new(big.Int).Sqrt(value.Num())
	d := new(big.Int).Sqrt(value.Denom())
	if new(big.Int).Mul(new(big.Int).Set(n), n).Cmp(value.Num()) != 0 || new(big.Int).Mul(new(big.Int).Set(d), d).Cmp(value.Denom()) != 0 {
		return nil
	}
	return new(big.Rat).SetFrac(n, d)
}
func nativeGeometryRationalBound(value *big.Rat) (*big.Rat, error) {
	if value != nil && (value.Num().BitLen() > nativeGeometryMaxRationalBits || value.Denom().BitLen() > nativeGeometryMaxRationalBits) {
		return nil, fmt.Errorf("geometry rational bit budget exceeded")
	}
	return value, nil
}
func (g nativeGeometryGuides) exactFormula(fields []string) (*big.Rat, error) {
	// Caller validates operator arity and numeric operands before this method.
	if len(fields) < 2 {
		return nil, nil
	}
	operands := make([]*big.Rat, len(fields)-1)
	for i, token := range fields[1:] {
		operands[i] = g.exactOperand(token)
	}
	x := operands[0]

	if fields[0] == "+-" {
		if operands[1] != nil && operands[1].Sign() == 0 && g.operandSymbol(fields[1]) == g.operandSymbol(fields[3]) {
			return big.NewRat(0, 1), nil
		}
		if x != nil && x.Sign() == 0 && g.operandSymbol(fields[2]) == g.operandSymbol(fields[3]) {
			return big.NewRat(0, 1), nil
		}
	}
	// Select an exact operand when inexact comparison bounds prove the branch.
	// This retains exact adjustment values pinned between transcendental limits.
	if fields[0] == "min" || fields[0] == "max" || fields[0] == "pin" {
		a, b := g.intervalOperand(fields[1]), g.intervalOperand(fields[2])
		if fields[0] == "min" {
			if a.hi <= b.lo {
				return operands[0], nil
			}
			if b.hi <= a.lo {
				return operands[1], nil
			}
		}
		if fields[0] == "max" {
			if a.lo >= b.hi {
				return operands[0], nil
			}
			if b.lo >= a.hi {
				return operands[1], nil
			}
		}
		if fields[0] == "pin" {
			c := g.intervalOperand(fields[3])
			if b.hi < a.lo {
				return operands[0], nil
			}
			if b.lo >= a.hi && b.hi <= c.lo {
				return operands[1], nil
			}
			if b.lo >= a.hi && b.lo > c.hi {
				return operands[2], nil
			}
		}
	}
	if fields[0] == "?:" {
		if x == nil {
			return nil, nil
		}
		if x.Sign() > 0 {
			return operands[1], nil
		}
		return operands[2], nil
	}
	for _, operand := range operands {
		if operand == nil {
			return nil, nil
		}
	}
	var y, z *big.Rat
	if len(operands) > 1 {
		y = operands[1]
	}
	if len(operands) > 2 {
		z = operands[2]
	}
	r := new(big.Rat)
	switch fields[0] {
	case "val":
		r.Set(x)
	case "+-":
		r.Sub(new(big.Rat).Add(x, y), z)
	case "*/":
		if z.Sign() == 0 {
			return nil, fmt.Errorf("zero geometry divisor")
		}
		r.Quo(new(big.Rat).Mul(x, y), z)
	case "+/":
		if z.Sign() == 0 {
			return nil, fmt.Errorf("zero geometry divisor")
		}
		r.Quo(new(big.Rat).Add(x, y), z)
	case "abs":
		r.Abs(x)
	case "min":
		if x.Cmp(y) < 0 {
			r.Set(x)
		} else {
			r.Set(y)
		}
	case "max":
		if x.Cmp(y) > 0 {
			r.Set(x)
		} else {
			r.Set(y)
		}
	case "pin":
		if y.Cmp(x) < 0 {
			r.Set(x)
		} else if y.Cmp(z) > 0 {
			r.Set(z)
		} else {
			r.Set(y)
		}
	case "sqrt":
		if x.Sign() < 0 {
			return nil, fmt.Errorf("negative geometry square root")
		}
		return nativeGeometryRationalBound(nativeGeometryExactSqrt(x))
	case "mod":
		r.Add(new(big.Rat).Mul(x, x), new(big.Rat).Mul(y, y))
		r.Add(r, new(big.Rat).Mul(z, z))
		return nativeGeometryRationalBound(nativeGeometryExactSqrt(r))
	case "cat2", "sat2":
		root := nativeGeometryExactSqrt(new(big.Rat).Add(new(big.Rat).Mul(y, y), new(big.Rat).Mul(z, z)))
		if root == nil {
			return nil, nil
		}
		if root.Sign() == 0 {
			return nil, fmt.Errorf("undefined geometry angle")
		}
		if fields[0] == "cat2" {
			r.Mul(x, y)
		} else {
			r.Mul(x, z)
		}
		r.Quo(r, root)
	case "at2":
		if x.Sign() == 0 && y.Sign() == 0 {
			return nil, fmt.Errorf("undefined geometry angle")
		}
		if x.Sign() == 0 {
			if y.Sign() > 0 {
				r.SetInt64(5400000)
			} else {
				r.SetInt64(-5400000)
			}
		} else if y.Sign() == 0 {
			if x.Sign() < 0 {
				r.SetInt64(10800000)
			} else {
				r.SetInt64(0)
			}
		} else {
			return nil, nil
		}
	case "sin", "cos", "tan":
		angle, _ := y.Float64()
		if !nativeGeometryFinite(angle) {
			return nil, fmt.Errorf("geometry angle exceeds precision bounds")
		}
		quarter := new(big.Rat).Quo(y, big.NewRat(5400000, 1))
		if !quarter.IsInt() {
			return nil, nil
		}
		cardinal := new(big.Int).Mod(quarter.Num(), big.NewInt(4)).Int64()
		if fields[0] == "tan" {
			if cardinal%2 != 0 {
				return nil, fmt.Errorf("undefined geometry tangent")
			}
			r.SetInt64(0)
		} else {
			sign := int64(0)
			if fields[0] == "sin" {
				if cardinal == 1 {
					sign = 1
				}
				if cardinal == 3 {
					sign = -1
				}
			} else {
				if cardinal == 0 {
					sign = 1
				}
				if cardinal == 2 {
					sign = -1
				}
			}
			r.Mul(x, big.NewRat(sign, 1))
		}
	default:
		return nil, nil
	}
	return nativeGeometryRationalBound(r)
}

func (g nativeGeometryGuides) operandSymbol(token string) string {
	if exact := g.exactOperand(token); exact != nil {
		return "r:" + exact.RatString()
	}
	if symbol, ok := g.symbols[token]; ok {
		return symbol
	}
	return "name:" + token
}
func (g nativeGeometryGuides) formulaSymbol(fields []string, exact *big.Rat) string {
	if exact != nil {
		return "r:" + exact.RatString()
	}
	if fields[0] == "val" {
		return g.operandSymbol(fields[1])
	}
	if fields[0] == "+-" {
		y, z := g.exactOperand(fields[2]), g.exactOperand(fields[3])
		if y != nil && z != nil && y.Cmp(z) == 0 {
			return g.operandSymbol(fields[1])
		}
		x := g.exactOperand(fields[1])
		if x != nil && z != nil && x.Cmp(z) == 0 {
			return g.operandSymbol(fields[2])
		}
	}
	parts := []string{fields[0]}
	for _, token := range fields[1:] {
		parts = append(parts, g.operandSymbol(token))
	}
	return fmt.Sprintf("expr:%x", sha256.Sum256([]byte(strings.Join(parts, "\x00"))))
}
