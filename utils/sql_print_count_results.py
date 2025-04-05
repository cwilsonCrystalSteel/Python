# -*- coding: utf-8 -*-
"""
Created on Sun Mar 30 13:06:47 2025

@author: Netadmin
"""

from sqlalchemy import text

def table_exists(engine, schema, table):
    with engine.connect() as connection:
        check_query = text(
            "SELECT EXISTS ("
            "SELECT 1 FROM information_schema.tables "
            "WHERE table_schema = :schema AND table_name = :table)"
        )
        result = connection.execute(check_query, {"schema": schema, "table": table})
        return result.scalar()


def print_count_results(engine, schema, table, suffix_text):
    if table_exists(engine, schema, table):
        with engine.connect() as connection:
            result = connection.execute(text(f"select count(*) from {schema}.{table}"))
            for row in result:
                continue
        
        print(f"There are {row[0]} rows in {schema}.{table} {suffix_text}")
    else:
        print(f"Table does not exist: {schema}.{table}")
        
def print_count_results_where(engine, schema, table, fieldname, equal_to_value, suffix_text):
    if table_exists(engine, schema, table):
        with engine.connect() as connection:
            result = connection.execute(text(f"select count(*) from {schema}.{table} where {fieldname}={equal_to_value}"))
            for row in result:
                continue
        
        print(f"There are {row[0]} rows in {schema}.{table} with {fieldname}={equal_to_value} {suffix_text}")
    else:
        print(f"Table does not exist: {schema}.{table}")
