import pytest
from dataclasses import dataclass
from typing import List
import tempfile
import json
from datetime import datetime

from SeatingPlanner import (
    Attendee, TableConstraints, SeatingOptimizer, SeatingHistory
)


@pytest.fixture
def sample_attendees():
    """Fixture providing a diverse set of attendees for testing"""
    return [
        Attendee("Person1", "F", "senior", "Sales", assign_head_table=True),
        Attendee("Person2", "M", "junior", "Engineering"),
        Attendee("Person3", "F", "senior", "Marketing"),
        Attendee("Person4", "M", "junior", "Sales"),
        Attendee("Person5", "F", "senior", "Engineering"),
        Attendee("Person6", "M", "junior", "Marketing"),
        Attendee("Person7", "F", "senior", "Sales"),
        Attendee("Person8", "M", "junior", "Engineering"),
        Attendee("Person9", "F", "senior", "Marketing"),
        Attendee("Person10", "M", "senior", "Sales", assign_head_table=True),
        Attendee("Person11", "F", "junior", "Engineering"),
        Attendee("Person12", "M", "senior", "Marketing"),
        Attendee("Person13", "F", "junior", "Sales"),
        Attendee("Person14", "M", "senior", "Engineering"),
        Attendee("Person15", "F", "junior", "Marketing"),
        Attendee("Person16", "M", "senior", "Sales"),
        Attendee("Person17", "F", "junior", "Engineering"),
        Attendee("Person18", "M", "senior", "Marketing"),
        Attendee("Person19", "F", "junior", "Sales"),
        Attendee("Person20", "M", "senior", "Engineering"),
        Attendee("Person21", "F", "junior", "Marketing"),
        Attendee("Person22", "M", "senior", "Sales"),
    ]


@pytest.fixture
def history_file():
    """Fixture providing a temporary history file"""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        json.dump([], f)
        return f.name


def test_table_constraints_validation():
    """Test that TableConstraints validates inputs correctly"""
    # Valid constraints
    constraints = TableConstraints(min_seats=4, max_seats=8)
    assert constraints.min_seats == 4
    assert constraints.max_seats == 8

    # Invalid constraints
    with pytest.raises(ValueError):
        TableConstraints(min_seats=8, max_seats=4)  # min > max

    with pytest.raises(ValueError):
        TableConstraints(min_seats=1, max_seats=8)  # min < 2


def test_table_sizes(sample_attendees):
    """Test that tables respect the configured size constraints"""
    test_sizes = [(4, 8)]  # (min_seats, max_seats) combinations

    for min_size, max_size in test_sizes:
        constraints = TableConstraints(min_seats=min_size, max_seats=max_size)
        optimizer = SeatingOptimizer(constraints=constraints)

        # Try with different numbers of tables
        for num_tables in [3, 4]:  # Including head table
            arrangement = optimizer.optimize_seating(sample_attendees, num_tables=num_tables)

            # Check head table size
            assert len(arrangement[0]) == max_size, \
                f"Head table size {len(arrangement[0])} doesn't match max_size {max_size}"

            # Check other table sizes
            for i, table in enumerate(arrangement[1:], 1):
                assert min_size <= len(table) <= max_size, \
                    f"Table {i} size {len(table)} outside range [{min_size}, {max_size}]"


def test_head_table_assignments(sample_attendees):
    """Test that head table assignments are respected"""
    constraints = TableConstraints(min_seats=4, max_seats=8)
    optimizer = SeatingOptimizer(constraints=constraints)

    arrangement = optimizer.optimize_seating(sample_attendees, num_tables=3)
    head_table = arrangement[0]

    # Check that pre-assigned attendees are at head table
    pre_assigned = [a for a in sample_attendees if a.assign_head_table]
    for attendee in pre_assigned:
        assert any(a.name == attendee.name for a in head_table), \
            f"Pre-assigned attendee {attendee.name} not at head table"


def test_table_distribution(sample_attendees):
    """Test that attendees are distributed reasonably across tables"""
    constraints = TableConstraints(min_seats=4, max_seats=8)
    optimizer = SeatingOptimizer(constraints=constraints)

    num_tables = 3
    arrangement = optimizer.optimize_seating(sample_attendees, num_tables=num_tables)

    # Check total attendees
    total_seated = sum(len(table) for table in arrangement)
    assert total_seated == len(sample_attendees), \
        f"Total seated {total_seated} doesn't match attendees {len(sample_attendees)}"

    # Check no duplicates
    all_seated = [a.name for table in arrangement for a in table]
    assert len(all_seated) == len(set(all_seated)), "Duplicate attendees found"


def test_insufficient_attendees():
    """Test handling of insufficient attendees"""
    constraints = TableConstraints(min_seats=4, max_seats=8)
    optimizer = SeatingOptimizer(constraints=constraints)

    # Create small attendee list
    few_attendees = [
        Attendee(f"Person{i}", "F", "senior", "Sales")
        for i in range(5)  # Not enough for two tables with min_seats=4
    ]

    with pytest.raises(ValueError) as exc_info:
        optimizer.optimize_seating(few_attendees, num_tables=2)
    assert "Need at least 12 attendees: 8 for head table and 4 each for 1 other tables" in str(exc_info.value)


def test_seating_history(sample_attendees, history_file):
    """Test that seating history affects future arrangements"""
    constraints = TableConstraints(min_seats=4, max_seats=8)
    history = SeatingHistory(filename=history_file)
    optimizer = SeatingOptimizer(constraints=constraints, history=history)

    # Generate and save first arrangement
    arrangement1 = optimizer.optimize_seating(sample_attendees, num_tables=3)
    history.save_arrangement(arrangement1)

    # Generate second arrangement
    arrangement2 = optimizer.optimize_seating(sample_attendees, num_tables=3)

    # Check that arrangements are different (history should influence seating)
    def get_table_pairs(arrangement):
        pairs = set()
        for table in arrangement:
            for i, a1 in enumerate(table):
                for a2 in table[i + 1:]:
                    pairs.add(tuple(sorted([a1.name, a2.name])))
        return pairs

    pairs1 = get_table_pairs(arrangement1)
    pairs2 = get_table_pairs(arrangement2)

    # Some pairs should be different due to history penalty
    assert pairs1 != pairs2, "Seating arrangements identical despite history"


def test_variable_table_sizes(sample_attendees):
    """Detailed test of table size behavior with different configurations"""
    for max_size in [6, 8, 10]:
        constraints = TableConstraints(min_seats=4, max_seats=max_size)
        optimizer = SeatingOptimizer(constraints=constraints)

        arrangement = optimizer.optimize_seating(sample_attendees, num_tables=4)

        # Print detailed size information for debugging
        print(f"\nTest with max_size={max_size}:")
        for i, table in enumerate(arrangement):
            print(f"Table {i} size: {len(table)}")

        # Verify head table size
        assert len(arrangement[0]) == max_size, \
            f"Head table has {len(arrangement[0])} seats instead of {max_size}"

        # Verify other table sizes
        other_table_sizes = [len(table) for table in arrangement[1:]]
        for size in other_table_sizes:
            assert 4 <= size <= max_size, \
                f"Found table with {size} seats, outside range [4, {max_size}]"


if __name__ == '__main__':
    pytest.main([__file__])