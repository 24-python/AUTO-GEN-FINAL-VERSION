"""
CRUD операции для ручного управления данными
"""

from sqlalchemy.orm import Session
from typing import List, Optional
from . import models


# === КАТЕГОРИИ ===
def get_categories(db: Session) -> List[models.Category]:
    return db.query(models.Category).order_by(models.Category.sort_order).all()


def get_category(db: Session, category_id: int) -> Optional[models.Category]:
    return db.query(models.Category).filter(models.Category.id == category_id).first()


def create_category(db: Session, name: str, sort_order: int = 0) -> models.Category:
    category = models.Category(name=name, sort_order=sort_order)
    db.add(category)
    db.commit()
    db.refresh(category)
    return category


def update_category(db: Session, category_id: int, **kwargs) -> Optional[models.Category]:
    category = get_category(db, category_id)
    if category:
        for key, value in kwargs.items():
            if hasattr(category, key):
                setattr(category, key, value)
        db.commit()
        db.refresh(category)
    return category


def delete_category(db: Session, category_id: int) -> bool:
    category = get_category(db, category_id)
    if category:
        db.delete(category)
        db.commit()
        return True
    return False


# === ОБЪЕКТЫ ===
def get_objects(db: Session, category_id: int = None) -> List[models.Object]:
    query = db.query(models.Object)
    if category_id:
        query = query.filter(models.Object.category_id == category_id)
    return query.order_by(models.Object.sort_priority).all()


def get_object(db: Session, object_id: int) -> Optional[models.Object]:
    return db.query(models.Object).filter(models.Object.id == object_id).first()


def get_object_by_name(db: Session, name: str) -> Optional[models.Object]:
    return db.query(models.Object).filter(models.Object.name == name).first()


def create_object(db: Session, name: str, base_name: str, category_id: int,
                  modifier: str = None, sort_priority: int = 0) -> models.Object:
    obj = models.Object(
        name=name,
        base_name=base_name,
        modifier=modifier,
        category_id=category_id,
        sort_priority=sort_priority
    )
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


def update_object(db: Session, object_id: int, **kwargs) -> Optional[models.Object]:
    obj = get_object(db, object_id)
    if obj:
        for key, value in kwargs.items():
            if hasattr(obj, key):
                setattr(obj, key, value)
        db.commit()
        db.refresh(obj)
    return obj


def delete_object(db: Session, object_id: int) -> bool:
    obj = get_object(db, object_id)
    if obj:
        db.delete(obj)
        db.commit()
        return True
    return False


# === ИНСТРУКЦИИ ===
def get_instructions(db: Session, object_id: int = None) -> List[models.Instruction]:
    query = db.query(models.Instruction)
    if object_id:
        query = query.filter(models.Instruction.object_id == object_id)
    return query.all()


def get_instruction(db: Session, instruction_id: int) -> Optional[models.Instruction]:
    return db.query(models.Instruction).filter(models.Instruction.id == instruction_id).first()


def create_instruction(db: Session, object_id: int, **kwargs) -> models.Instruction:
    instruction = models.Instruction(object_id=object_id, **kwargs)
    db.add(instruction)
    db.commit()
    db.refresh(instruction)
    return instruction


def update_instruction(db: Session, instruction_id: int, **kwargs) -> Optional[models.Instruction]:
    instruction = get_instruction(db, instruction_id)
    if instruction:
        for key, value in kwargs.items():
            if hasattr(instruction, key):
                setattr(instruction, key, value)
        db.commit()
        db.refresh(instruction)
    return instruction


def delete_instruction(db: Session, instruction_id: int) -> bool:
    instruction = get_instruction(db, instruction_id)
    if instruction:
        db.delete(instruction)
        db.commit()
        return True
    return False