#---------------------
# ある集合内に含まれる要素をグループに分割するためのスクリプト
# 何人かのグループに適当に分割したいときに使用する
#---------------------

# メンバーを保持するクラス
class Members
	@members = Array.new()

	def getSplitMembers(num)
		@retAry = Array.new()

		@members.shuffle!

		@members.length.times do |i|
			if i%num == 0 then
				@retAry.push("----------------")
			end
			@retAry.push(@members.at(i))
		end
		@retAry.push("----------------")

		return @retAry
	end

	def setMembers(members)
		@members = members
	end
end

# メンバーの登録(ここにメンバーを追加)
ary = Array["a","b","c","d","e","f","g","h","i","j","k","l","m","n"]

# 何人グループを作るか？
groupNum = 3

# メンバークラスの生成
members = Members.new()

# メンバークラスに名前をセット
members.setMembers(ary)

# 指定人数でグルーピングされた全メンバーを抽出
resultMembers = members.getSplitMembers(groupNum)

# 出力
puts(resultMembers)
