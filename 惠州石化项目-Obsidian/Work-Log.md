# Work-Log

## 2026-06-03

- 创建项目 Obsidian 资料库初始结构。
- 记录项目目标、模板、测试数据范围和当前核心问题。
- 记录用户提供的 2026 年 5 月“其他在途”截图线索。
- 已按规则读取 `AGENTS.md` 和 `Project-Brief.md`。
- 定位“其他在途”读取逻辑：`code/intransit_read.py` 负责 2-1/2-2 在途数量和人民币金额；`code/intransit_price_pack.py` 负责 2-3 在途美元/桶价格包。
- 发现根因线索：代码原本用精确页签名判断 `其他在途`，而 2026 年 4 月暂估表 `2026年4月原油暂估新.xlsx` 的实际页签为 `其他在途 `（末尾空格），会导致整张表跳过。5 月截图也可能存在同类不可见尾部空格。
- 已修复：两个读取入口均增加页签名首尾空格兼容，找不到精确页签时按 `strip()` 后匹配真实页签。
- 验证结果：3 月暂估表可读出 1 条 DES/DAT/DAP；4 月 `其他在途 ` 修复后可读出 2 条 DES/DAT/DAP，且价格包可读出 DES/DAT/DAP 合计美元/桶。
- 语法检查：`intransit_read.py` 和 `intransit_price_pack.py` 通过 `py_compile`。
- 下一步：拿到 2026 年 5 月原油暂估源文件后复测；同时建议增加更明确的缺页签/空数据日志。
- 为当前代码建立本地 Git 备份：在项目根目录初始化仓库，提交 `code/`、`.gitignore` 和 `惠州石化项目-Obsidian/`，排除月份 Excel 数据、PDF、Office 临时锁文件和 `__pycache__`。
- 本地备份提交：`16112bd`，提交信息 `Backup Huizhou cost report project`。
- 上传远端状态：本机未安装 GitHub CLI `gh`，项目此前也没有远端；GitHub 连接器可访问账号 `itxuqiqi`，但未发现明确对应本项目的仓库。等待用户指定目标仓库。
- 用户指定远端仓库：`https://github.com/itxuqiqi/Laptop_Project.git`。
- 已将本地 `main` 推送到 GitHub 远端 `origin/main`。推送时本机全局 Git 将 HTTPS 改写为 SSH，因此补写了 GitHub 官方 SSH host key 到 `known_hosts` 后完成推送。
