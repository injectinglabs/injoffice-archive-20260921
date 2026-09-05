module github.com/injectinglabs/injoffice/go/officecompat

go 1.23.0

require (
	github.com/injectinglabs/injoffice/go/docxpatch v0.0.0
	github.com/injectinglabs/injoffice/go/pptxpatch v0.0.0
	github.com/injectinglabs/injoffice/go/xlsxpatch v0.0.0
)

replace github.com/injectinglabs/injoffice/go/docxpatch => ../docxpatch

replace github.com/injectinglabs/injoffice/go/pptxpatch => ../pptxpatch

replace github.com/injectinglabs/injoffice/go/xlsxpatch => ../xlsxpatch
