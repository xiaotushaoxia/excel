package excel

import (
	"fmt"
)

func ToCell(c string, r int) string {
	return fmt.Sprintf("%s%d", c, r)
}

// TitleToNumber https://leetcode.cn/problems/excel-sheet-column-number/  官方题解
func TitleToNumber(columnTitle string) (number int) {
	for i, multiple := len(columnTitle)-1, 1; i >= 0; i-- {
		k := columnTitle[i] - 'A' + 1
		number += int(k) * multiple
		multiple *= 26
	}
	return
}

// NumberToTitle https://leetcode.cn/problems/excel-sheet-column-title/ 官方题解
func NumberToTitle(columnNumber int) string {
	var ans []byte
	for columnNumber > 0 {
		columnNumber--
		ans = append(ans, 'A'+byte(columnNumber%26))
		columnNumber /= 26
	}
	for i, n := 0, len(ans); i < n/2; i++ {
		ans[i], ans[n-1-i] = ans[n-1-i], ans[i]
	}
	return string(ans)
}
