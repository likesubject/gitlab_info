# -- coding: utf-8 --
from tools import Gitlab, Excel
import click


@click.command()
@click.option("--token", default='_vTxX9faeFzc4WRFjWdu', help="Gitlab private token")
@click.option("--service_url", default='http://git.zhaoqi.info:9998', help="Gitlab service url")
@click.option("--api_version", default='4', help="Gitlab api version")
def run(token, service_url, api_version):
    gitlab = Gitlab(token, service_url, api_version)
    with Excel.context() as excel:
        for project in gitlab.projects:
            excel.write(project)


if __name__ == '__main__':
    # run('_vTxX9faeFzc4WRFjWdu', 'http://git.zhaoqi.info:9998', '4') #test
    try:
        run()
    except ValueError as e:
        print(e)
