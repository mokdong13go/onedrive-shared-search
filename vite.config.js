import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// GitHub Pages 는 https://<사용자>.github.io/<저장소>/ 하위 경로로 서비스되므로
// 빌드 시 base 를 저장소 이름으로 맞춘다. 로컬 dev 에서는 "/" 를 쓴다.
// 저장소 이름을 바꾸면 아래 값도 함께 바꿀 것.
const REPO_NAME = "onedrive-shared-search";

export default defineConfig(({ command }) => ({
  base: command === "build" ? `/${REPO_NAME}/` : "/",
  plugins: [react()],
  server: {
    port: 3000,
  },
}));
