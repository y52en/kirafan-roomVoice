import "./style.css";

import Handsontable from "handsontable";
import ExcelJS from "exceljs";
import "handsontable/dist/handsontable.full.css";

// const app = document.querySelector<HTMLDivElement>("#app")!;

// app.innerHTML = `
//   <h1>Hello Vite!</h1>
//   <a href="https://vitejs.dev/guide/features.html" target="_blank">Documentation</a>
// `;

(async () => {
  type tweet = {
    m_ID: number;
    m_LimitedCharaId: number[];
    m_NamedType: number;
    m_Text: string;
    m_VoiceCueName: string;
  };
  type room = {
    m_DBAccessID: number;
    m_Name: string;
    m_ObjEventArgs: number[];
  };
  const NamedList_prom = fetch(
    "https://database.kirafan.cn/database/NamedList.json"
  ).then((r) => r.json());
  const RoomObjectList_prom: Promise<room[]> = fetch(
    "https://database.kirafan.cn/database/RoomObjectList.json"
  ).then((r) => r.json());
  const TweetList_prom: Promise<tweet[]> = fetch(
    "https://database.kirafan.cn/database/TweetList.json"
  ).then((r) => r.json());

  const [NamedList, RoomObjectList, TweetList] = await Promise.all([
    NamedList_prom,
    RoomObjectList_prom,
    TweetList_prom,
  ]);

  const find_chara = (id: Number): string => {
    return NamedList.find(
      (chara: { m_NamedType: number; m_FullName: string }) =>
        chara.m_NamedType === id
    ).m_FullName;
  };

  // m_ID :8桁ルーム, 先頭21 平日,22 土日
  const db = TweetList.filter((x: tweet) => {
    return (
      x.m_ID.toString().length === 9 ||
      x.m_ID.toString().match(/^(21|22)\d{3}$/)
    );
  }).map((tweet: tweet) => {
    const voice_type: [number, string] = (() => {
      const id_str = tweet.m_ID.toString();
      const id_len = id_str.length;

      if (id_len === 9) {
        return [0, "特定の家具"];
      } else if (id_str.match(/^21\d{3}$/)) {
        return [1, "平日"];
      } else if (id_str.match(/^22\d{3}$/)) {
        return [2, "土日"];
      } else {
        throw new Error("unknown type");
      }
    })();
    return [
      find_chara(tweet.m_NamedType),
      voice_type[1],
      tweet.m_Text,
      RoomObjectList.filter((x) => x.m_ObjEventArgs[1] === tweet.m_ID)
        .map((x) => x.m_Name)
        .join("\n"),
      tweet.m_LimitedCharaId.length > 0 ? "限定キャラ" : "",
    ];
  });
  const data_headers = [
    "キャラ名",
    "方法",
    "セリフ",
    "対応する家具",
    "限定キャラのみのセリフ",
  ];

  const data = [
    // data_headers,
    ...db,
  ];

  const container = document.getElementById("table") as HTMLDivElement;
  if (!container) {
    throw new Error("container not found");
  }
  document
    .querySelectorAll(".loading")
    .forEach((el) => el.classList.add("display-none"));
  document
    .querySelectorAll(".loaded")
    .forEach((el) => el.classList.remove("display-none"));

  new Handsontable(container, {
    data: data,
    colHeaders: data_headers,
    columns: [
      { type: "text" },
      { type: "text" },
      { type: "text" },
      { type: "text" },
      { type: "text" },
    ],
    colWidths: [100, 100, 200, 200, 150],
    rowHeaders: true,
    // colHeaders: true,
    width: "auto",
    filters: true,
    dropdownMenu: [
      "filter_by_condition",
      "filter_action_bar",
      "filter_by_value",
    ],
    contextMenu: true,
    manualRowMove: true,
    licenseKey: "non-commercial-and-evaluation",
  });

  document.querySelectorAll(".download_xlsx").forEach(async (elm) =>
    elm.addEventListener("click", async (_e) => {
      const workbook = new ExcelJS.Workbook();
      workbook.addWorksheet("sheet1");
      const worksheet = workbook.getWorksheet("sheet1");

      // ["キャラ名", "方法", "セリフ", "対応する家具", "限定キャラのみのセリフ"]
      worksheet.columns = [
        { header: "キャラ名", key: "0" },
        { header: "方法", key: "1" },
        { header: "セリフ", key: "2" },
        { header: "対応する家具", key: "3" },
        // { header: "セリフ", key: "4" },
        { header: "限定キャラのみのセリフ", key: "4" },
      ];

      const rows = db
        .map((row, i) => {
          if (i === 0) return;
          const output: { [s: string]: any } = {};
          row.forEach((x, i) => (output[String(i)] = x));

          return output;
        })
        .filter((x) => x);

      worksheet.addRows(rows);
      worksheet.autoFilter = "A1:E" + rows.length;

      const uint8Array = await workbook.xlsx.writeBuffer();
      const blob = new Blob([uint8Array], { type: "application/octet-binary" });
      fileDownload_fromURL(blob, "RoomVoice.xlsx");
    })
  );
})();

function fileDownload_fromURL(blob: Blob, fileName = "img.png") {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}
