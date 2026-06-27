如果你想为本代码库做出贡献，请遵循以下步骤：

1. Fork 仓库并创建你的分支：
  ```pwsh
  git checkout -b new-feature
  uv sync --group dev
  ...
  uvx ty check .  # 检查类型，确保通过
  ```

2. 提交你的更改，使用清晰的提交信息，遵循[约定式提交](https://www.conventionalcommits.org/)规范：
   ```pwsh
   git commit -m "feat: add some feature"
   ```

3. 推送到你的分支：
   ```pwsh
   git push origin new-feature
   ```

4. 创建一个新的 Pull Request，格式参考 `../../.github/PULL_REQUEST_TEMPLATE.md`。
