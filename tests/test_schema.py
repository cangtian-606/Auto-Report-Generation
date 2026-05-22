"""SchemaValidator unit tests"""

import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from src.processing.schema import DataSchema, SchemaValidator, SheetSchema, FieldSpec


class TestDictValidation:
    def test_required_field_present_passes(self):
        schema = DataSchema(sheets={
            "全局": SheetSchema(fields={"name": FieldSpec(required=True)})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"全局": {"name": "test"}})
        assert len(errors) == 0

    def test_required_field_missing_reports_error(self):
        schema = DataSchema(sheets={
            "全局": SheetSchema(fields={"name": FieldSpec(required=True)})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"全局": {"other": "val"}})
        assert len(errors) == 1
        assert "全局.name" in errors[0]

    def test_type_check_string(self):
        schema = DataSchema(sheets={
            "全局": SheetSchema(fields={"name": FieldSpec(field_type="str")})
        })
        validator = SchemaValidator(schema)
        errors = validator.validate({"全局": {"name": 123}})
        assert len(errors) == 1


class TestDataFrameValidation:
    def test_dataframe_required_column_present_passes(self):
        schema = DataSchema(sheets={
            "商品": SheetSchema(
                required=True,
                fields={"名称": FieldSpec(required=True)}
            )
        })
        df = pd.DataFrame({"名称": ["A", "B"], "金额": [100, 200]})
        validator = SchemaValidator(schema)
        errors = validator.validate({"商品": df})
        assert len(errors) == 0

    def test_dataframe_column_missing_reports_error(self):
        schema = DataSchema(sheets={
            "商品": SheetSchema(
                required=True,
                fields={"金额": FieldSpec(required=True)}
            )
        })
        df = pd.DataFrame({"名称": ["A", "B"]})
        validator = SchemaValidator(schema)
        errors = validator.validate({"商品": df})
        assert len(errors) == 1
        assert "商品" in errors[0]

    def test_dataframe_rows_present_passes(self):
        schema = DataSchema(sheets={
            "商品": SheetSchema(required=True)
        })
        df = pd.DataFrame({"名称": ["A"]})
        validator = SchemaValidator(schema)
        errors = validator.validate({"商品": df})
        assert len(errors) == 0

    def test_dataframe_empty_reports_error_when_required(self):
        schema = DataSchema(sheets={
            "商品": SheetSchema(required=True)
        })
        df = pd.DataFrame()
        validator = SchemaValidator(schema)
        errors = validator.validate({"商品": df})
        assert len(errors) == 1


class TestDataSchemaConstruction:
    def test_from_dict_simple(self):
        d = {"sheets": {"全局": {"required": True, "fields": {"name": True}}}}
        schema = DataSchema.from_dict(d)
        assert schema.sheets["全局"].required is True
        assert schema.sheets["全局"].fields["name"].required is True
