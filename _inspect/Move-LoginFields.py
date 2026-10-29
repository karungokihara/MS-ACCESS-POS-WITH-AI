p = r"C:\Users\user\Downloads\LLC\_inspect\F_Login_now.txt"
text = open(p, encoding="utf-16").read().splitlines()
start = None
end = None
for i, line in enumerate(text):
    if 'Name ="txtUsername"' in line and start is None:
        start = i - 20
    if 'Name ="EmptyCell34"' in line:
        end = i + 30
print("range", start, end)
skip = {
    "GroupTable =2",
    "LayoutGroup =2",
    "ColumnStart =1",
    "ColumnEnd =1",
    "RowStart =1",
    "RowEnd =1",
    "RowStart =2",
    "RowEnd =2",
}
out = []
for i, line in enumerate(text):
    s = line.strip()
    if start is not None and start <= i <= end:
        if s in skip:
            continue
        if s == "Left =4365":
            line = line.replace("4365", "8460")
        if s == "LayoutCachedLeft =4365":
            line = line.replace("4365", "8460")
        if s == "LayoutCachedWidth =9465":
            line = line.replace("9465", "13560")
    out.append(line)
open(p, "w", encoding="utf-16").write("\r\n".join(out) + "\r\n")
print("rewritten", len(out))
