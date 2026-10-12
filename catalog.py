"""Muc luc SGK Toan 8, Ket noi tri thuc voi cuoc song.
Doi chieu tu anh muc luc tap mot PDF trang 5, tap hai PDF trang 4.
So trang la so IN tren sach, khong phai chi so trang PDF.
"""
CHAPTERS = [
 (1,1,"Đa thức",[(1,"Đơn thức",5),(2,"Đa thức",11),(3,"Phép cộng và phép trừ đa thức",15),(4,"Phép nhân đa thức",19),(5,"Phép chia đa thức cho đơn thức",22)]),
 (2,1,"Hằng đẳng thức đáng nhớ và ứng dụng",[(6,"Hiệu hai bình phương. Bình phương của một tổng hay một hiệu",29),(7,"Lập phương của một tổng. Lập phương của một hiệu",34),(8,"Tổng và hiệu hai lập phương",37),(9,"Phân tích đa thức thành nhân tử",42)]),
 (3,1,"Tứ giác",[(10,"Tứ giác",48),(11,"Hình thang cân",52),(12,"Hình bình hành",57),(13,"Hình chữ nhật",64),(14,"Hình thoi và hình vuông",67)]),
 (4,1,"Định lí Thalès",[(15,"Định lí Thalès trong tam giác",76),(16,"Đường trung bình của tam giác",81),(17,"Tính chất đường phân giác của tam giác",84)]),
 (5,1,"Dữ liệu và biểu đồ",[(18,"Thu thập và phân loại dữ liệu",90),(19,"Biểu diễn dữ liệu bằng bảng, biểu đồ",93),(20,"Phân tích số liệu thống kê dựa vào biểu đồ",99)]),
 (6,2,"Phân thức đại số",[(21,"Phân thức đại số",4),(22,"Tính chất cơ bản của phân thức đại số",8),(23,"Phép cộng và phép trừ phân thức đại số",15),(24,"Phép nhân và phép chia phân thức đại số",20)]),
 (7,2,"Phương trình bậc nhất và hàm số bậc nhất",[(25,"Phương trình bậc nhất một ẩn",27),(26,"Giải bài toán bằng cách lập phương trình",33),(27,"Khái niệm hàm số và đồ thị của hàm số",40),(28,"Hàm số bậc nhất và đồ thị của hàm số bậc nhất",47),(29,"Hệ số góc của đường thẳng",51)]),
 (8,2,"Mở đầu về tính xác suất của biến cố",[(30,"Kết quả có thể và kết quả thuận lợi",59),(31,"Cách tính xác suất của biến cố bằng tỉ số",63),(32,"Mối liên hệ giữa xác suất thực nghiệm với xác suất và ứng dụng",67)]),
 (9,2,"Tam giác đồng dạng",[(33,"Hai tam giác đồng dạng",78),(34,"Ba trường hợp đồng dạng của hai tam giác",83),(35,"Định lí Pythagore và ứng dụng",93),(36,"Các trường hợp đồng dạng của hai tam giác vuông",98),(37,"Hình đồng dạng",104)]),
 (10,2,"Một số hình khối trong thực tiễn",[(38,"Hình chóp tam giác đều",112),(39,"Hình chóp tứ giác đều",117)]),
]
PRACTICE = [
 (1, "Công thức lãi kép",111),(1,"Thực hiện tính toán trên đa thức với phần mềm GeoGebra",113),
 (1,"Vẽ hình đơn giản với phần mềm GeoGebra",115),(1,"Phân tích đặc điểm khí hậu Việt Nam",120),
 (2,"Một vài ứng dụng của hàm số bậc nhất trong tài chính",125),
 (2,"Ứng dụng định lí Thalès, định lí Pythagore và tam giác đồng dạng để đo chiều cao, khoảng cách",128),
 (2,"Thực hành tính toán trên phân thức đại số và vẽ đồ thị hàm số với phần mềm GeoGebra",130),
 (2,"Mô tả thí nghiệm ngẫu nhiên với phần mềm Excel",133),
]
# Các đoạn ghi chú dưới đây được biên soạn lại từ phần khoanh kiến thức trong
# một số trang đã xem trực tiếp; KHÔNG phải toàn văn SGK.
VERIFIED_NOTES = {
 1: {"focus":"Đơn thức, đơn thức thu gọn, hệ số, phần biến, bậc của đơn thức; đơn thức đồng dạng.","formulas":["Đơn thức đồng dạng: có cùng phần biến (sau khi thu gọn).", "Cộng/trừ đơn thức đồng dạng: cộng/trừ các hệ số, giữ nguyên phần biến."],"page_ref":"Tập 1, tr. 5–10"},
 2: {"focus":"Khái niệm đa thức nhiều biến, hạng tử, đa thức thu gọn và bậc của đa thức.","formulas":["Đa thức thu gọn: không còn các hạng tử đồng dạng.","Bậc của đa thức thu gọn khác 0 là bậc cao nhất của các hạng tử."],"page_ref":"Tập 1, tr. 11–14"},
 3: {"focus":"Phép cộng, phép trừ đa thức; quy tắc bỏ dấu ngoặc và thu gọn hạng tử đồng dạng.","formulas":["A − B = A + (−B)."],"page_ref":"Tập 1, tr. 15–16"},
 4: {"focus":"Nhân đơn thức với đơn thức, đơn thức với đa thức và đa thức với đa thức.","formulas":["A(B + C) = AB + AC."],"page_ref":"Tập 1, tr. 19–21"},
 5: {"focus":"Chia đơn thức cho đơn thức và chia đa thức cho đơn thức (trường hợp chia hết).","formulas":["(A + B) : M = A : M + B : M, khi các phép chia đều thực hiện được."],"page_ref":"Tập 1, tr. 22–24"},
 6: {"focus":"Hiệu hai bình phương và bình phương một tổng, một hiệu.","formulas":["A² − B² = (A − B)(A + B)","(A + B)² = A² + 2AB + B²","(A − B)² = A² − 2AB + B²"],"page_ref":"Tập 1, tr. 29–33"},
 7: {"focus":"Lập phương một tổng, lập phương một hiệu.","formulas":["(A + B)³ = A³ + 3A²B + 3AB² + B³","(A − B)³ = A³ − 3A²B + 3AB² − B³"],"page_ref":"Tập 1, tr. 34–36"},
 8: {"focus":"Hai hằng đẳng thức tổng và hiệu hai lập phương.","formulas":["A³ + B³ = (A + B)(A² − AB + B²)","A³ − B³ = (A − B)(A² + AB + B²)"],"page_ref":"Tập 1, tr. 37–39"},
 9: {"focus":"Phân tích đa thức thành nhân tử: đặt nhân tử chung, dùng hằng đẳng thức, nhóm hạng tử.","formulas":["AB + AC = A(B + C)"],"page_ref":"Tập 1, tr. 42–44"},
 10:{"focus":"Khái niệm tứ giác, tứ giác lồi, đỉnh, cạnh, góc và đường chéo.","formulas":["Tổng bốn góc trong một tứ giác bằng 360°."],"page_ref":"Tập 1, tr. 48–51"},
 11:{"focus":"Hình thang cân, tính chất về cạnh bên, đường chéo và dấu hiệu nhận biết.","formulas":["Trong hình thang cân, hai cạnh bên bằng nhau và hai đường chéo bằng nhau."],"page_ref":"Tập 1, tr. 52–55"},
 12:{"focus":"Hình bình hành, tính chất và dấu hiệu nhận biết.","formulas":["Trong hình bình hành, hai cạnh đối bằng nhau, hai góc đối bằng nhau và hai đường chéo cắt nhau tại trung điểm mỗi đường."],"page_ref":"Tập 1, tr. 57–61"},
 13:{"focus":"Hình chữ nhật, tính chất hai đường chéo và dấu hiệu nhận biết.","formulas":["Trong hình chữ nhật, hai đường chéo bằng nhau và cắt nhau tại trung điểm mỗi đường."],"page_ref":"Tập 1, tr. 64–66"},
 14:{"focus":"Hình thoi, hình vuông và tính chất, dấu hiệu nhận biết.","formulas":["Hình thoi là tứ giác có bốn cạnh bằng nhau."],"page_ref":"Tập 1, tr. 67–72"},
 21:{"focus":"Khái niệm phân thức đại số, hai phân thức bằng nhau, điều kiện xác định và giá trị.","formulas":["Phân thức A/B có B là đa thức khác đa thức 0; giá trị xác định khi B ≠ 0."],"page_ref":"Tập 2, tr. 4–7"},
 22:{"focus":"Tính chất cơ bản của phân thức; rút gọn và quy đồng mẫu thức.","formulas":["A/B = (A·M)/(B·M) khi M là đa thức khác 0.","A/B = (A:N)/(B:N) khi N là nhân tử chung khác 0."],"page_ref":"Tập 2, tr. 8–12"},
 23:{"focus":"Cộng và trừ các phân thức cùng mẫu, khác mẫu và nhiều phân thức.","formulas":["A/M + B/M = (A + B)/M","A/M − B/M = (A − B)/M"],"page_ref":"Tập 2, tr. 15–19"},
 24:{"focus":"Nhân, chia phân thức đại số và vận dụng tính chất phép toán.","formulas":["(A/B)·(C/D) = AC/BD","(A/B):(C/D) = (A/B)·(D/C) với phân thức chia khác 0."],"page_ref":"Tập 2, tr. 20–22"},
 25:{"focus":"Phương trình một ẩn, nghiệm của phương trình, phương trình bậc nhất một ẩn và cách giải.","formulas":["ax + b = 0 (a ≠ 0) có nghiệm x = −b/a."],"page_ref":"Tập 2, tr. 27–32"},
 26:{"focus":"Giải bài toán bằng cách lập phương trình: chọn ẩn, lập phương trình, giải và đối chiếu điều kiện.","formulas":["Đặt ẩn và điều kiện → lập phương trình → giải → đối chiếu, kết luận."],"page_ref":"Tập 2, tr. 33–36"},
 27:{"focus":"Khái niệm hàm số, mặt phẳng tọa độ, tọa độ điểm và đồ thị hàm số.","formulas":["Trong mặt phẳng tọa độ, điểm M(x₀; y₀) có hoành độ x₀, tung độ y₀."],"page_ref":"Tập 2, tr. 40–46"},
 28:{"focus":"Hàm số bậc nhất, bảng giá trị và đồ thị hàm số bậc nhất.","formulas":["y = ax + b (a ≠ 0) là hàm số bậc nhất.","Đồ thị y = ax + b (a ≠ 0) là một đường thẳng."],"page_ref":"Tập 2, tr. 47–50"},
 29:{"focus":"Hệ số góc của đường thẳng và quan hệ giữa hai đường thẳng.","formulas":["Đường thẳng y = ax + b (a ≠ 0) có hệ số góc là a."],"page_ref":"Tập 2, tr. 51–54"},
 30:{"focus":"Liệt kê kết quả có thể của hành động/thí nghiệm và kết quả thuận lợi cho biến cố.","formulas":["Kết quả thuận lợi cho biến cố là kết quả làm biến cố đó xảy ra."],"page_ref":"Tập 2, tr. 59–62"},
 31:{"focus":"Tính xác suất của biến cố bằng tỉ số trong trường hợp các kết quả đồng khả năng.","formulas":["P(E) = số kết quả thuận lợi cho E / tổng số kết quả có thể (đồng khả năng)."],"page_ref":"Tập 2, tr. 63–66"},
 32:{"focus":"Xác suất thực nghiệm và mối liên hệ với xác suất; ước lượng từ thí nghiệm.","formulas":["Xác suất thực nghiệm của E = số lần E xảy ra / tổng số lần thực hiện thí nghiệm."],"page_ref":"Tập 2, tr. 67–73"},
}

def catalog_json():
 out=[]
 for no,volume,title,lessons in CHAPTERS:
  out.append({"id":no,"volume":volume,"name":title,"lessons":[{"id":num,"name":name,"printed_page":pg,"verified_summary":VERIFIED_NOTES.get(num)} for num,name,pg in lessons],"sections":["Luyện tập chung","Bài tập cuối chương"]})
 return {"chapters":out,"practice":[{"volume":v,"title":t,"printed_page":p} for v,t,p in PRACTICE],"source_note":"Danh mục được đối chiếu theo mục lục 2 tập SGK do giáo viên cung cấp. Ghi chú lý thuyết là bản tóm tắt có giới hạn, không thay thế sách giáo khoa."}
