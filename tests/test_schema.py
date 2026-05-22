"""SchemaValidator unit tests"""

import sys
from pathlib import Path

import pytest

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from src.processing.schema import DataSchema, SchemaValidator, EntrySchema, FieldSpec


class TestDictValidation:
    def test_required_field_present_passes(self):
        schema = DataSchema(entries={
            "全局": EntrySchema(fields={"name": FieldSpec(required=True)})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"全局": {"name": "test"}})
        assert len(errors) == 0

    def test_required_field_missing_reports_error(self):
        schema = DataSchema(entries={
            "全局": EntrySchema(fields={"name": FieldSpec(required=True)})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"全局": {"other": "val"}})
        assert len(errors) == 1
        assert "全局.name" in errors[0]

    def test_type_check_string(self):
        schema = DataSchema(entries={
            "全局": EntrySchema(fields={"name": FieldSpec(field_type="str")})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"全局": {"name": 123}})
        assert len(errors) == 1


class TestListValidation:
    def test_list_required_fields_pass(self):
        schema = DataSchema(entries={
            "股东": EntrySchema(fields={"名称": FieldSpec(required=True)})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"股东": [{"名称": "张伟"}, {"名称": "李芳"}]})
        assert len(errors) == 0

    def test_list_field_missing_reports_error(self):
        schema = DataSchema(entries={
            "股东": EntrySchema(fields={"认缴额": FieldSpec(required=True)})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"股东": [{"名称": "张伟"}]})
        assert len(errors) == 1
        assert "股东[0].认缴额" in errors[0]

    def test_list_min_items_pass(self):
        schema = DataSchema(entries={
            "股东": EntrySchema(min_items=2)
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"股东": [{"名称": "A"}, {"名称": "B"}]})
        assert len(errors) == 0

    def test_list_min_items_fail(self):
        schema = DataSchema(entries={
            "股东": EntrySchema(min_items=3)
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"股东": [{"名称": "A"}]})
        assert len(errors) == 1
        assert "至少需要 3 条" in errors[0]

    def test_empty_list_required_reports_error(self):
        schema = DataSchema(entries={
            "股东": EntrySchema(required=True)
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"股东": []})
        assert len(errors) == 1
        assert "无数据" in errors[0]

    def test_list_type_check(self):
        schema = DataSchema(entries={
            "股东": EntrySchema(fields={"金额": FieldSpec(field_type="int")})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"股东": [{"金额": "不是数字"}]})
        assert len(errors) == 1
        assert "类型错误" in errors[0]


class TestDataSchemaConstruction:
    def test_from_dict_simple(self):
        d = {"entries": {"全局": {"required": True, "fields": {"name": True}}}}
        schema = DataSchema.from_dict(d)
        assert schema.entries["全局"].required is True
        assert schema.entries["全局"].fields["name"].required is True

    def test_from_dict_with_list_entry(self):
        d = {"entries": {"股东": {"min_items": 2, "fields": {"名称": True}}}}
        schema = DataSchema.from_dict(d)
        assert schema.entries["股东"].min_items == 2
        assert schema.entries["股东"].fields["名称"].required is True

    def test_from_dict_backward_compat_sheets(self):
        d = {"sheets": {"全局": {"required": True, "fields": {"name": True}}}}
        schema = DataSchema.from_dict(d)
        assert "全局" in schema.entries
        assert schema.entries["全局"].required is True
