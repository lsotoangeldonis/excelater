"""tests/test_scheduler.py — Tests de lógica del scheduler"""
import pytest
from apscheduler.triggers.cron import CronTrigger

from app.database import ScheduleType
from app.scheduler import build_trigger


def test_build_trigger_once_daily():
    trigger = build_trigger(ScheduleType.ONCE_DAILY, {"hour": 9, "minute": 30})
    assert isinstance(trigger, CronTrigger)
    fields = {f.name: str(f) for f in trigger.fields}
    assert fields["hour"] == "9"
    assert fields["minute"] == "30"


def test_build_trigger_interval_hours():
    trigger = build_trigger(ScheduleType.INTERVAL, {"hours": 2, "start_hour": 6, "start_minute": 0})
    assert isinstance(trigger, CronTrigger)
    fields = {f.name: str(f) for f in trigger.fields}
    assert "2" in fields["hour"]


def test_build_trigger_interval_minutes():
    trigger = build_trigger(ScheduleType.INTERVAL, {"minutes": 30, "start_hour": 8})
    assert isinstance(trigger, CronTrigger)
    fields = {f.name: str(f) for f in trigger.fields}
    assert "30" in fields["minute"]


def test_build_trigger_cron():
    trigger = build_trigger(ScheduleType.CRON, {"cron": "0 6 * * 1-5"})
    assert isinstance(trigger, CronTrigger)
    fields = {f.name: str(f) for f in trigger.fields}
    assert fields["hour"] == "6"
    assert fields["minute"] == "0"


def test_build_trigger_unknown_type():
    with pytest.raises(ValueError, match="desconocido"):
        build_trigger("invalid_type", {})
