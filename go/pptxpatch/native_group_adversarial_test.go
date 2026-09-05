package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestNativeGroupAffineComponentsEnforceExactSafeIntegerBoundaries(t *testing.T) {
	t.Parallel()

	transform := func(x, y, cx, cy int64) NativeTransform {
		return NativeTransform{X: &x, Y: &y, Cx: &cx, Cy: &cy}
	}
	tests := []struct {
		name                         string
		outer, child                 NativeTransform
		wantX, wantY, wantTX, wantTY int64
		wantErr                      bool
	}{
		{
			name:  "asymmetric negative coordinate spaces",
			outer: transform(-7, 11, 3, 5), child: transform(-2, 4, 2, 5),
			wantX: 1_500_000, wantY: 1_000_000, wantTX: -4, wantTY: 7,
		},
		{
			name:  "exact maximum safe coefficient",
			outer: transform(0, 0, nativeMaxSafeInteger, nativeMaxSafeInteger),
			child: transform(0, 0, nativeGroupTransformPPM, nativeGroupTransformPPM),
			wantX: nativeMaxSafeInteger, wantY: nativeMaxSafeInteger,
		},
		{
			name:  "fractional x translation",
			outer: transform(0, 0, 3, 1), child: transform(1, 0, 2, 1),
			wantErr: true,
		},
		{
			name:  "coefficient exceeds safe integer",
			outer: transform(0, 0, nativeMaxSafeInteger, 1), child: transform(0, 0, 1, 1),
			wantErr: true,
		},
		{
			name:    "translation exceeds safe integer",
			outer:   transform(nativeMaxSafeInteger, 0, nativeMaxSafeInteger, 1),
			child:   transform(-nativeGroupTransformPPM, 0, nativeGroupTransformPPM, 1),
			wantErr: true,
		},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			x, y, tx, ty, err := nativeGroupAffineComponents(test.outer, test.child)
			if test.wantErr {
				if err == nil {
					t.Fatalf("unsafe or inexact affine was accepted: (%d, %d, %d, %d)", x, y, tx, ty)
				}
				return
			}
			if err != nil || x != test.wantX || y != test.wantY || tx != test.wantTX || ty != test.wantTY {
				t.Fatalf("affine mismatch: got (%d, %d, %d, %d, %v), want (%d, %d, %d, %d)", x, y, tx, ty, err, test.wantX, test.wantY, test.wantTX, test.wantTY)
			}
		})
	}
}

func TestExtractNativePPTXGroupTransactionLifecycleIsFailClosed(t *testing.T) {
	t.Parallel()

	presentationNS, drawingNS := nativeGroupTestNamespaces(false)
	group := nativeNestedGroupXML(presentationNS, drawingNS, "")
	group = strings.Replace(group,
		`<p:cNvPr id="5" name="Nested Rect"/>`,
		`<p:cNvPr id="5" name="Nested Rect"><a:extLst><a:ext uri="preserve"><x:metadata xmlns:x="urn:injoffice:test"/></a:ext></a:extLst></p:cNvPr>`, 1)
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
	}})

	for _, test := range []struct {
		name       string
		beginErr   bool
		nilTx      bool
		issueErr   bool
		commitErr  bool
		wantCommit int
		wantRoll   int
		wantOK     bool
	}{
		{name: "begin error", beginErr: true},
		{name: "nil transaction", nilTx: true},
		{name: "issue error rolls back", issueErr: true, wantRoll: 1},
		{name: "commit error rolls back", commitErr: true, wantCommit: 1, wantRoll: 1},
		{name: "success commits without rollback", wantCommit: 1, wantOK: true},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			issuer := &nativeAdversarialAtomicIssuer{
				beginErr: test.beginErr, nilTx: test.nilTx,
				tx: &nativeAdversarialAtomicTransaction{issueErr: test.issueErr, commitErr: test.commitErr},
			}
			issuer.tx.issuer = issuer
			deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: issuer})
			if test.wantOK {
				if err != nil || len(deck.Slides) != 1 {
					t.Fatalf("successful transaction failed: deck=%#v err=%v", deck, err)
				}
			} else if err == nil || len(deck.Slides) != 0 {
				t.Fatalf("failed transaction returned a deck: deck=%#v err=%v", deck, err)
			}
			if issuer.directIssues != 0 || issuer.tx.commits != test.wantCommit || issuer.tx.rollbacks != test.wantRoll {
				t.Fatalf("transaction lifecycle mismatch: direct=%d commits=%d rollbacks=%d", issuer.directIssues, issuer.tx.commits, issuer.tx.rollbacks)
			}
		})
	}
}

type nativeAdversarialAtomicIssuer struct {
	beginErr     bool
	nilTx        bool
	directIssues int
	tx           *nativeAdversarialAtomicTransaction
}

func (issuer *nativeAdversarialAtomicIssuer) IssueNativePassthroughToken(NativePassthroughTokenRequest) (string, error) {
	issuer.directIssues++
	return "", fmt.Errorf("direct issuance is forbidden")
}

func (issuer *nativeAdversarialAtomicIssuer) BeginNativePassthroughTokenTransaction() (NativePassthroughTokenTransaction, error) {
	if issuer.beginErr {
		return nil, fmt.Errorf("injected begin failure")
	}
	if issuer.nilTx {
		return nil, nil
	}
	return issuer.tx, nil
}

type nativeAdversarialAtomicTransaction struct {
	issuer    *nativeAdversarialAtomicIssuer
	issueErr  bool
	commitErr bool
	issues    int
	commits   int
	rollbacks int
}

func (transaction *nativeAdversarialAtomicTransaction) IssueNativePassthroughToken(request NativePassthroughTokenRequest) (string, error) {
	transaction.issues++
	if transaction.issueErr {
		return "", fmt.Errorf("injected issue failure")
	}
	return "adversarial-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
}

func (transaction *nativeAdversarialAtomicTransaction) CommitNativePassthroughTokens() error {
	transaction.commits++
	if transaction.commitErr {
		return fmt.Errorf("injected commit failure")
	}
	return nil
}

func (transaction *nativeAdversarialAtomicTransaction) RollbackNativePassthroughTokens() {
	transaction.rollbacks++
}
