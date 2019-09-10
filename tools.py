# -- coding: utf-8 --
from __future__ import unicode_literals

import logging
import time
from contextlib import contextmanager
from gitlab.v4.objects import *
import gitlab
import xlwt


class Gitlab(object):
    def __init__(self, private_token, service_url, api_version='3'):
        self.service_url = service_url
        self.api_version = api_version
        self.gl = gitlab.Gitlab(
            service_url, private_token=private_token, api_version=str(api_version)
        )
        try:
            self._projects = self.gl.projects.list()
        except Exception as e:
            logging.error(repr(e))
            raise ValueError('Gitlab connect error')
        except gitlab.exceptions.GitlabGetError:
            raise ValueError('Gitlab project get error')
        if self._projects is None:
            ValueError('project not be found')

    @property
    def projects(self):
        for project in self._projects:
            yield GitlabProject(project, self)


class GitlabProject(object):
    """
    GitlabProject
    """

    """
    属性列表
    [(属性名, 属性中文名, 属性方法)]
    """
    _attr_list = [
        ('repo_name', '仓库名称', 'get_repo_name'),
        ('created_on', '创建时间', 'get_created_on'),
        ('repo_url', '仓库地址', 'get_repo_url'),
        ('description', '仓库描述', 'get_description'),
        ('master_users', '仓库管理员', 'get_repo_master_users'),
        ('work_group', '工作组', 'get_work_group')
    ]

    def __init__(self, project, gitlab):
        self._project = project
        self.gitlab = gitlab

    def get_repo_name(self):
        return self._project.name

    def get_repo_url(self):
        return self._project.web_url

    def get_description(self):
        return self._project.description

    def get_created_on(self):
        return self._project.created_at

    def format_members(self, members):
        str_users = ''
        point = ''
        for member in members:
            str_users += point + self._get_user_name(member)
            point = ', '
        return str_users

    def _get_user_name(self, member):
        if self.gitlab.api_version == '3':
            return None
        else:
            return member.name

    def get_repo_master_users(self):
        master_members = []
        if self.gitlab.api_version == '3':
            pass
        else:
            members = self._project.members.list()
            for member in members:
                if member.access_level == gitlab.MASTER_ACCESS:
                    master_members.append(member)
        return self.format_members(master_members)

    def get_work_group(self):
        if self.gitlab.api_version == '3':
            return self._project.namespace_id
        else:
            return self._project.namespace.get('full_path', None)

    @property
    def attrs(self):
        return self._attr_list


class Excel(object):
    def __init__(self):
        self.sheet = self.work_book.add_sheet('Gitlab Info')
        self.setting_field = False
        self.row_count = 0

    @classmethod
    @contextmanager
    def context(cls, file_name_prefix='gitlab-info'):
        cls.work_book = xlwt.Workbook()
        inst = cls()
        yield inst
        str_time = time.strftime(str("%Y-%m-%d-%H-%M-%S"), time.localtime())
        file_name = '{0}-{1}.xls'.format(file_name_prefix, str_time)
        inst.work_book.save(file_name)

    def _init_fields(self, project):
        """初始化字段"""
        for i, attr in enumerate(project.attrs):
            self.sheet.write(0, i, self.get_column_name(attr[0], attr[1]))
        self.row_count += 1

    def write(self, project):
        if not self.setting_field:
            self.setting_field = True
            self._init_fields(project)
        for i, attr in enumerate(project.attrs):
            self.sheet.write(self.row_count, i, getattr(project, attr[2])())
        self.row_count += 1

    @staticmethod
    def get_column_name(name, title):
        return '{1}({0})'.format(name, title)

