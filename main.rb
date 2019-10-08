# coding: utf-8
require 'rubyXL'
require 'rubyXL/convenience_methods'

def convert_alphabet_to_col_num(alphabet)
  ret = 0
  size = alphabet.size
  alphabet.to_s.each_char do |char|
    ret += (char.ord - 'A'.ord + 1) * 26 ** (size - 1)
    size -= 1
  end
  ret
end

p convert_alphabet_to_col_num('XEC')

return


book = RubyXL::Parser.parse('t2.xlsx')

# original from http://secret-garden.hatenablog.com/entry/2017/09/21/174348

book.each do |sheet|
  sheet.each { |row|
    row && row.cells.each_with_index { |cell, i|
      val = cell && cell.value
     if val.nil?
       print "#{i}[nil]"
     else
       p val
     end
    }
  }
end

return


src  = "test.xlsx"
dest = "output.xlsx"

# エクセルファイルの読み込み
book = RubyXL::Parser.parse(src)





# シートを取得
sheet = book[0]

# シートのダンプ1

sheet.each { |row|
   row && row.cells.each { |cell|
     val = cell && cell.value
     p val
   }
}

# index指定シート
(0..2).each do |x|
  (0..2).each do |y|
    p sheet[x][y]
  end
end


# シートのセルを結合する

# 行3, 列 0 のセル位置
cell1 = [3, 0]
# 行3, 列 1 のセル位置
cell2 = [3, 1]

# cell1 と cell2 を結合する
sheet.merge_cells *cell1, *cell2
# 以下と等価
# sheet.merge_cells 3, 0, 3, 1

# ファイルの保存
book.write(dest)
