const form = document.querySelector("#compare-form");
const status = document.querySelector("#status");
const button = form.querySelector("button");

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(form);
  button.disabled = true;
  status.textContent = "엑셀을 비교하는 중입니다…";
  try {
    const response = await fetch("/api/compare", { method: "POST", body: data });
    if (!response.ok) throw new Error((await response.json()).error || "처리에 실패했습니다.");
    const blob = await response.blob();
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "B_색상표시_비교결과.xlsx";
    link.click();
    URL.revokeObjectURL(link.href);
    const stats = JSON.parse(decodeURIComponent(response.headers.get("X-Compare-Stats") || "%7B%7D"));
    status.textContent = `완료: 노랑 ${stats.yellowCells ?? 0}셀 · 빨강 ${stats.redCells ?? 0}셀`;
  } catch (error) {
    status.textContent = error.message;
  } finally {
    button.disabled = false;
  }
});
