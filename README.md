# TY Equipment Wiki

网站资料由 `data/TY_Product_Database.xlsx` 管理。日常更新不需要修改 `index.html`。

## 更新产品资料

1. 在 GitHub 打开 `data/TY_Product_Database.xlsx`。
2. 点击下载并使用 Microsoft Excel 打开。
3. 只修改以下四张工作表：
   - `Ex-stock`
   - `Specifications`
   - `Product Guide`
   - `Brand Lineup`
4. 不要修改工作表名称和第一行的栏目名称。
5. 保存文件，文件名必须保持为 `TY_Product_Database.xlsx`。
6. 回到 GitHub 的 `data` 文件夹。
7. 点击 **Add file → Upload files**。
8. 上传新的 Excel，覆盖旧文件。
9. 点击 **Commit changes**。
10. 等待 Vercel 部署完成后刷新网站。

## 上传产品图片

1. 图片放在 GitHub 的 `product-images` 文件夹。
2. `Ex-stock` 工作表的 `Image` 栏只填写文件名，例如 `SRR-LV681-SG.png`。
3. Excel 中的文件名必须与 GitHub 图片完全相同，包括大小写和 `.png`、`.jpg`。
4. 没有图片时可以把 `Image` 留空，网站会显示 placeholder。

## 各工作表的用途

### Ex-stock

控制库存搜索结果、数量、价格和产品图片。每个 SKU 只能出现一次。价格直接填写数字，例如 `2480`，不需要输入 `$`。

### Specifications

控制点击产品卡后出现的 specification popup。只有在这里出现的 SKU 才能打开 specification。

### Product Guide

控制 Product Guide 的产品、品牌、focus、lead time、questions 和 important notes。同一格有多个项目时，使用 Excel 的换行输入。

`Images` 的填写格式：

```text
guide-upright.png | Upright refrigerator examples
guide-undercounter.png | Undercounter refrigerator examples
```

### Brand Lineup

控制点击品牌 Logo 后出现的 SKU。每个 SKU 使用一行。如果该 SKU 同时存在于 Ex-stock 和 Specifications，点击时会打开 specification popup。

## 如果上传后网站没有更新

1. 打开 Vercel 项目的 **Deployments**。
2. 查看最新一项是否显示 **Ready**。
3. 如果显示 **Error**，打开错误记录。它会说明哪张工作表或哪一行有问题。
4. 修正 Excel 后重新上传。
5. 上一个正常的网站版本不会因为 Excel 错误而被覆盖。

## 不要修改

- 不要删除 `scripts` 文件夹。
- 不要修改 `package.json` 或 `vercel.json`。
- 不要随意修改 `index.html` 中的资料结构。
- 不要把 Excel 改成其他文件名。
