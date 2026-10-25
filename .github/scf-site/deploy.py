#!/usr/bin/env python3
"""把网站静态文件 + SCF handler 打包成 zip，通过腾讯云 API 更新函数代码。

用法（GitHub Action 或本地）：
  TENCENTCLOUD_SECRET_ID=xxx TENCENTCLOUD_SECRET_KEY=yyy python3 deploy.py
可选环境变量：SCF_SITE_FUNCTION（默认 social-practice-site）、SCF_REGION（默认 ap-shanghai）
"""
import base64
import io
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))          # .github/scf-site/
REPO_ROOT = os.path.normpath(os.path.join(HERE, '..', '..'))
HANDLER = os.path.join(HERE, 'handler.js')

# 网页实际需要的静态文件（仓库根目录下）
SITE_FILES = ['index.html', 'app.js', 'payout.js', 'fares.json', 'station_codes.json']


def build_zip() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(HANDLER, 'index.js')
        for f in SITE_FILES:
            src = os.path.join(REPO_ROOT, f)
            if not os.path.exists(src):
                sys.exit(f'缺少文件：{src}')
            z.write(src, 'site/' + f)
    return buf.getvalue()


def main() -> None:
    sid = os.environ.get('TENCENTCLOUD_SECRET_ID')
    skey = os.environ.get('TENCENTCLOUD_SECRET_KEY')
    fn = os.environ.get('SCF_SITE_FUNCTION', 'social-practice-site')
    region = os.environ.get('SCF_REGION', 'ap-shanghai')
    if not sid or not skey:
        sys.exit('缺少 TENCENTCLOUD_SECRET_ID / TENCENTCLOUD_SECRET_KEY 环境变量')

    from tencentcloud.common import credential
    from tencentcloud.scf.v20180416 import models, scf_client

    client = scf_client.ScfClient(credential.Credential(sid, skey), region)
    req = models.UpdateFunctionCodeRequest()
    req.FunctionName = fn
    req.ZipFile = base64.b64encode(build_zip()).decode()
    client.UpdateFunctionCode(req)
    print(f'已部署到函数 {fn}（{region}），含 {len(SITE_FILES)} 个静态文件')


if __name__ == '__main__':
    main()
