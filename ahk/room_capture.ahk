#Requires AutoHotkey v2.0

F8::
{
    t := WinGetTitle("A")
    room := t

    ; (일단 테스트를 위해 '카카오톡' 체크를 제거했습니다)
    ; 나중에 안정화할 때 다시 넣거나, 더 정확한 조건으로 바꿀 수 있어요.

    ; 제목 정리
    room := RegExReplace(room, "\s*-\s*카카오톡\s*$", "")
    room := RegExReplace(room, "카카오톡", "")
    room := Trim(room)

    ; 폴더/파일에 못 쓰는 문자 제거
    room := RegExReplace(room, "[\\/:*?`"<>|]", "_")
    if (room = "")
        room := "알수없음"

    ctxPath := A_Temp "\kakao_room_ctx.txt"
    try FileDelete(ctxPath)
    FileAppend(room "|" A_Now, ctxPath, "UTF-8")

    ToolTip("캡처: " room, 20, 20)
    SetTimer(() => ToolTip(), -1500)
}
