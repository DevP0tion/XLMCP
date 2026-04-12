import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as insertImage } from "./insert-image.tool.js";
import { register as listImages } from "./list-images.tool.js";
import { register as manageImage } from "./manage-image.tool.js";

export function registerImageTools(server: McpServer) {
  insertImage(server);
  listImages(server);
  manageImage(server);
}
