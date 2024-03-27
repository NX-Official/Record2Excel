package Record2Excel

import (
	"github.com/xuri/excelize/v2"
)

type options struct {
	headerStyle     *excelize.Style
	contentStyle    *excelize.Style
	customStyleFunc func(record any) (style *excelize.Style)
}

var (
	defaultHeaderStyle  = StyleHorizontalCenter
	defaultContentStyle = StyleCenter
)

type Option interface {
	apply(*options)
}

type optionFunc func(*options)

func (f optionFunc) apply(opt *options) {
	f(opt)
}

func WithHeaderStyle(style *excelize.Style) Option {
	return optionFunc(func(o *options) {
		o.headerStyle = style
	})
}

func WithContentStyle(style *excelize.Style) Option {
	return optionFunc(func(o *options) {
		o.contentStyle = style
	})
}

func WithCustomStyleFunc(f func(record any) (style *excelize.Style)) Option {
	return optionFunc(func(o *options) {
		o.customStyleFunc = f
	})
}
