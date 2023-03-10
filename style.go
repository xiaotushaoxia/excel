package excel

type FormatStyle struct {
	Border        []Border    `json:"border,omitempty"`
	Fill          *Fill       `json:"fill,omitempty"`
	Font          *FormatFont `json:"font,omitempty"`
	Alignment     *Alignment  `json:"alignment,omitempty"`
	Protection    *Protection `json:"protection,omitempty"`
	NumFmt        int         `json:"number_format,omitempty"`
	DecimalPlaces int         `json:"decimal_places,omitempty"`
	CustomNumFmt  string      `json:"custom_number_format,omitempty"`
	Lang          string      `json:"lang,omitempty"`
	NegRed        bool        `json:"negred,omitempty"`
}

type Border struct {
	Type  string `json:"type,omitempty"`
	Color string `json:"color,omitempty"`
	Style int    `json:"style,omitempty"`
}

type Alignment struct {
	Horizontal      string `json:"horizontal,omitempty"`
	Indent          int    `json:"indent,omitempty"`
	JustifyLastLine bool   `json:"justify_last_line,omitempty"`
	ReadingOrder    uint64 `json:"reading_order,omitempty"`
	RelativeIndent  int    `json:"relative_indent,omitempty"`
	ShrinkToFit     bool   `json:"shrink_to_fit,omitempty"`
	TextRotation    int    `json:"text_rotation,omitempty"`
	Vertical        string `json:"vertical,omitempty"`
	WrapText        bool   `json:"wrap_text,omitempty"`
}

type Protection struct {
	Hidden bool `json:"hidden,omitempty"`
	Locked bool `json:"locked,omitempty"`
}

type Fill struct {
	Type    string   `json:"type,omitempty"`
	Pattern int      `json:"pattern,omitempty"`
	Color   []string `json:"color,omitempty"`
	Shading int      `json:"shading,omitempty"`
}

type FormatFont struct {
	Bold      bool   `json:"bold,omitempty"`
	Italic    bool   `json:"italic,omitempty"`
	Underline string `json:"underline,omitempty"`
	Family    string `json:"family,omitempty"`
	Size      int    `json:"size,omitempty"`
	Color     string `json:"color,omitempty"`
}
