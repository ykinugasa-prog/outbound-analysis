import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "アウトバウンド接続分析アプリ",
  description: "曜日・時間帯別の発信数、接続数、接続率を集計します。",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  );
}
