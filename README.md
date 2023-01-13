这就是个对excelize的简单封装
因为实际使用中有些接口用的不太舒服  包装一下顺手一点

1. 因为我项目中有对sheet的操作，所以先封了个FileSheet. 这样不用每次都 File.XXX(sheetName, xxx...), 而是FileSheet.XXX(xxx...)
2. excelize设置style居然要 NewStyle(`{"custom_number_format": "[$-380A]dddd\\,\\ dd\" de \"mmmm\" de \"yyyy;@"}`)。我很不理解，为什么要这样做。所以把源码里面的style复制出来包装了一下。包装后的用法见test，我感觉是更好的。然后很奇怪，网上找的一些例子，是有Style暴露出来的，不知道为什么我这里没有，只能用json字符串来NewStyle，不太懂，不管了
3. TitleToNumber就是 A->1, AA->27,  NumberToTitle就是 1->A, 27->AA， excelize里面也有类似的，但是转换是A->0, 0-A。 虽然我很喜欢0作为第一个元素，但是excel是1开头的，所以遵守excel让我比较适应（个人喜好）
