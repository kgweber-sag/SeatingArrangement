import itertools
import json
import random
from dataclasses import dataclass
from datetime import datetime
from typing import List, Dict, Tuple, Optional

import pandas as pd


@dataclass
class Attendee:
    name: str
    gender: str
    seniority: str
    field: str
    assign_head_table: bool = False

    def __hash__(self):
        return hash(self.name)


@dataclass
class TableConstraints:
    min_seats: int
    max_seats: int

    def __post_init__(self):
        if self.min_seats > self.max_seats:
            raise ValueError("Minimum seats cannot be greater than maximum seats")
        if self.min_seats < 2:
            raise ValueError("Minimum seats must be at least 2")



class SeatingHistory:
    def __init__(self, filename: Optional[str] = None, memory_events: int = 3):
        self.filename=filename
        self.memory_events = memory_events
        self.history = self.load_history(filename) if filename else []

    def load_history(self, filename: str) -> List[Dict]:
        """Load history from Excel file"""
        try:
            # Read Excel without setting index
            df = pd.read_excel(filename)

            # Convert Excel format to internal format for optimization
            events = []
            for column in df.columns:
                if column == 'Name':  # Skip the Name column
                    continue

                try:
                    try:
                        date = datetime.strptime(column, '%m/%d/%Y')
                    except ValueError:
                        date = column
                    # Find maximum table number in this column
                    max_table = 1  # Start with at least 1 table
                    for cell in df[column]:
                        if pd.isna(cell) or cell == '(did not attend)':
                            continue
                        if isinstance(cell, str) and '[head table]' in cell:
                            table_num = 1
                        else:
                            table_num = int(float(cell))
                        max_table = max(max_table, table_num)

                    event_data = {'date': date, 'arrangement': [[] for _ in range(max_table + 1)]}

                    # Group attendees by table number
                    for _, row in df.iterrows():
                        name = row['Name']
                        table_info = row[column]

                        if pd.isna(table_info) or table_info == '(did not attend)':
                            continue

                        # Parse table number and head table designation
                        if isinstance(table_info, str) and '[head table]' in table_info:
                            table_num = 0  # Head table is always index 0
                            is_head_table = True
                        else:
                            table_num = int(float(table_info)) - 1  # Convert to 0-based index
                            is_head_table = False

                        event_data['arrangement'][table_num].append({
                            'name': name,
                            'assign_head_table': is_head_table
                        })

                    events.append(event_data)
                except (ValueError, TypeError) as e:
                    continue  # Skip columns that aren't dates or have invalid data

            return sorted(events, key=lambda x: x['date'])

        except FileNotFoundError:
            return []

    def create_updated_history(self, arrangement: List[List['Attendee']],
                               previous_history: Optional[pd.DataFrame] = None,
                               event_date: Optional[datetime] = None) -> pd.DataFrame:
        """Create updated history DataFrame with new arrangement"""
        if event_date is None:
            event_date = datetime.today()

        column_name = event_date.strftime('%m/%d/%Y')

        # Start with previous history if provided
        if previous_history is not None:
            df = previous_history.copy()
        else:
            # Create new DataFrame with Name column
            df = pd.DataFrame(columns=['Name'])

        # Collect all attendee names
        current_names = {attendee.name for table in arrangement for attendee in table}
        existing_names = set(df['Name']) if 'Name' in df.columns else set()
        all_names = current_names.union(existing_names)

        # Ensure Name column exists and contains all names
        if 'Name' not in df.columns:
            df['Name'] = list(all_names)
        else:
            # Add any new names
            new_names = all_names - existing_names
            if new_names:
                new_rows = pd.DataFrame({'Name': list(new_names)})
                df = pd.concat([df, new_rows], ignore_index=True)

        # Create new column data
        new_data = {}
        for name in all_names:
            # Find attendee in current arrangement
            found = False
            for table_idx, table in enumerate(arrangement):
                for attendee in table:
                    if attendee.name == name:
                        if table_idx == 0:  # Head table
                            new_data[name] = f"1 [head table]" if attendee.assign_head_table else "1"
                        else:
                            new_data[name] = str(table_idx + 1)  # Convert to 1-based table numbers
                        found = True
                        break
                if found:
                    break
            if not found:
                new_data[name] = ''

        # Update DataFrame
        df[column_name] = df['Name'].map(new_data)

        # Sort columns: Name first, then dates
        date_columns = [col for col in df.columns if col != 'Name']

        # Convert all date strings to datetime for sorting
        def safe_date_parse(col):
            try:
                if isinstance(col, datetime):
                    return col
                return datetime.strptime(col, '%m/%d/%Y')
            except (ValueError, TypeError):
                return datetime.min  # Put non-date columns at the start

        date_columns.sort(key=safe_date_parse)
        df = df[['Name'] + date_columns]

        # Ensure all date columns are formatted consistently
        for col in date_columns:
            if isinstance(col, datetime):
                # Rename column to formatted string
                df.rename(columns={col: col.strftime('%m/%d/%Y')}, inplace=True)

        return df

    def get_recent_pairings(self) -> Dict[Tuple[str, str], List[int]]:
        """Get recent seating history for optimization"""
        recent_events = self.history[-self.memory_events:] if self.history else []
        pairings = {}

        all_attendees = set()
        for event in recent_events:
            for table in event['arrangement']:
                all_attendees.update(person['name'] for person in table)

        for name1, name2 in itertools.combinations(all_attendees, 2):
            pair = tuple(sorted([name1, name2]))
            pairings[pair] = [0] * len(recent_events)

        for event_idx, event in enumerate(recent_events):
            for table in event['arrangement']:
                names = [person['name'] for person in table]
                for name1, name2 in itertools.combinations(names, 2):
                    pair = tuple(sorted([name1, name2]))
                    pairings[pair][event_idx] = 1

        return pairings


class SeatingOptimizer:
    def __init__(self, constraints: TableConstraints, history: SeatingHistory = None,
                 weights: Dict[str, float] = {"gender_weight": 1.0, "seniority_weight": 1.0, "field_weight": 1.0}):
        self.constraints = constraints
        self.history = history or SeatingHistory()
        self.weights = weights

    def calculate_table_diversity_score(self, table: List[Attendee],
                                        gender_weight=1,
                                        seniority_weight=1,
                                        field_weight=1) -> float:
        if not table or len(table) < self.constraints.min_seats:
            return 0.0

        gender_ratio = sum(1 for a in table if a.gender == 'F') / len(table)
        gender_score = 1 - abs(0.5 - gender_ratio)

        senior_ratio = sum(1 for a in table if a.seniority == 'senior') / len(table)
        seniority_score = 1 - abs(0.5 - senior_ratio)

        unique_fields = len(set(a.field for a in table))
        field_score = unique_fields / len(table)

        return (gender_score * gender_weight + seniority_score * seniority_weight + field_score * field_weight) / sum(
            [gender_weight, seniority_weight, field_weight])

    def calculate_time_weighted_penalty(self, table: List[Attendee],
                                        recent_pairings: Dict[Tuple[str, str], List[int]]) -> float:
        if not table or len(table) < self.constraints.min_seats:
            return 0.0

        total_penalty = 0
        num_pairs = 0
        time_weights = [1.0, 0.6, 0.3]

        for person1, person2 in itertools.combinations(table, 2):
            pair = tuple(sorted([person1.name, person2.name]))
            if pair in recent_pairings:
                pair_history = recent_pairings[pair]
                weighted_penalty = sum(w * h for w, h in zip(time_weights[:len(pair_history)], pair_history))
                total_penalty += weighted_penalty
                num_pairs += 1

        if num_pairs == 0:
            return 1.0

        avg_penalty = total_penalty / num_pairs
        return max(0, 1 - avg_penalty)

    def create_balanced_tables(self, attendees: List[Attendee], num_tables: int) -> List[List[Attendee]]:
        """Distribute attendees across tables while respecting min/max constraints"""
        if not attendees:
            return []

        # Calculate target size for even distribution
        total_attendees = len(attendees)
        target_size = total_attendees // num_tables

        # Verify we can create valid tables
        if target_size < self.constraints.min_seats:
            raise ValueError(
                f"Too few attendees ({total_attendees}) to create {num_tables} tables with minimum {self.constraints.min_seats} seats")

        # Initialize empty tables
        tables = [[] for _ in range(num_tables)]
        remaining_attendees = attendees.copy()

        # First pass: ensure minimum sizes
        for table in tables:
            for _ in range(self.constraints.min_seats):
                if remaining_attendees:
                    table.append(remaining_attendees.pop())

        # Second pass: distribute remaining attendees
        while remaining_attendees:
            # Find table with fewest attendees that isn't at max capacity
            valid_tables = [t for t in tables if len(t) < self.constraints.max_seats]
            if not valid_tables:
                raise ValueError("Cannot distribute remaining attendees while respecting maximum table size")

            target_table = min(valid_tables, key=len)
            target_table.append(remaining_attendees.pop())

        return tables
    def optimize_seating(self, attendees: List[Attendee], num_tables: int,
                         head_table_assignments: Optional[List[Attendee]] = None,
                         iterations: int = 1000) -> List[List[Attendee]]:
        recent_pairings = self.history.get_recent_pairings()
        best_arrangement = []
        best_score = -1

        # Handle head table assignments
        head_table_attendees = head_table_assignments or [a for a in attendees if a.assign_head_table]
        other_attendees = [a for a in attendees if a not in head_table_attendees]

        # Verify total attendees meets minimum requirements
        total_attendees = len(attendees)
        remaining_tables = num_tables - 1  # Excluding head table
        min_required = self.constraints.max_seats + (self.constraints.min_seats * remaining_tables)

        if total_attendees < min_required:
            raise ValueError(
                f"Need at least {min_required} attendees: {self.constraints.max_seats} for head table "
                f"and {self.constraints.min_seats} each for {remaining_tables} other tables"
            )

        # Verify head table constraints
        if len(head_table_attendees) > self.constraints.max_seats:
            raise ValueError(f"Too many head table assignments (maximum is {self.constraints.max_seats})")

        for _ in range(iterations):
            try:
                # Always fill head table to exactly max_seats
                current_head_table = head_table_attendees.copy()
                available_for_head = [a for a in other_attendees if not a.assign_head_table]

                if len(current_head_table) < self.constraints.max_seats:
                    random.shuffle(available_for_head)
                    seats_needed = self.constraints.max_seats - len(current_head_table)
                    head_table_fills = available_for_head[:seats_needed]
                    current_head_table.extend(head_table_fills)
                    available_for_head = [a for a in available_for_head if a not in head_table_fills]

                # Create other tables with remaining attendees
                remaining_attendees = [a for a in other_attendees if a not in current_head_table]
                random.shuffle(remaining_attendees)

                # Calculate target size for remaining tables
                total_remaining = len(remaining_attendees)
                base_size = total_remaining // remaining_tables
                extra = total_remaining % remaining_tables

                other_tables = []
                start_idx = 0

                for i in range(remaining_tables):
                    # Add one extra person to first 'extra' tables
                    table_size = base_size + (1 if i < extra else 0)

                    # Ensure table size meets constraints
                    if table_size < self.constraints.min_seats:
                        raise ValueError(f"Cannot create table with only {table_size} attendees")
                    if table_size > self.constraints.max_seats:
                        raise ValueError(f"Table size {table_size} exceeds maximum of {self.constraints.max_seats}")

                    end_idx = start_idx + table_size
                    other_tables.append(remaining_attendees[start_idx:end_idx])
                    start_idx = end_idx

                current_arrangement = [current_head_table] + other_tables

                # Score the arrangement
                diversity_scores = [
                    self.calculate_table_diversity_score(
                        table,
                        **self.weights
                    ) for table in current_arrangement
                ]
                recency_scores = [
                    self.calculate_time_weighted_penalty(table, recent_pairings)
                    for table in current_arrangement
                ]

                # Calculate size balance score (excluding head table)
                other_table_sizes = [len(table) for table in other_tables]
                size_variations = [abs(s1 - s2) for s1, s2 in itertools.combinations(other_table_sizes, 2)]
                size_balance_score = 1.0 if not size_variations else (
                        1.0 - (sum(size_variations) / (len(other_tables) * self.constraints.max_seats))
                )

                # Combine scores (60% diversity, 25% recency, 15% size balance)
                current_score = (
                        0.6 * sum(diversity_scores) / len(current_arrangement) +
                        0.25 * sum(recency_scores) / len(current_arrangement) +
                        0.15 * size_balance_score
                )

                if current_score > best_score:
                    best_score = current_score
                    best_arrangement = current_arrangement

            except ValueError:
                continue

        if not best_arrangement:
            raise ValueError("Could not find valid seating arrangement")

        # Final verification
        if len(best_arrangement[0]) != self.constraints.max_seats:
            raise ValueError(
                f"Head table has {len(best_arrangement[0])} seats instead of required {self.constraints.max_seats}")

        return best_arrangement


def print_seating_arrangement(arrangement: List[List[Attendee]], history: SeatingHistory = None):
    recent_pairings = history.get_recent_pairings() if history else None

    # Print head table
    print("\nHead Table (Fixed Size):")
    print("-" * 50)
    print(f"Number of guests: {len(arrangement[0])} (Maximum capacity)")
    for attendee in arrangement[0]:
        assigned_marker = "*" if attendee.assign_head_table else " "
        print(f"{attendee.name:15} | {attendee.gender:6} | {attendee.seniority:6} | {attendee.field} {assigned_marker}")

    # Print other tables
    for i, table in enumerate(arrangement[1:], 2):
        print(f"\nTable {i}:")
        print("-" * 50)
        print(f"Number of guests: {len(table)} (Flexible capacity)")
        for attendee in table:
            print(f"{attendee.name:15} | {attendee.gender:6} | {attendee.seniority:6} | {attendee.field}")

        if recent_pairings:
            print("\nRecent interactions at this table:")
            for person1, person2 in itertools.combinations(table, 2):
                pair = tuple(sorted([person1.name, person2.name]))
                if pair in recent_pairings:
                    history = recent_pairings[pair]
                    recent_count = sum(history)
                    if recent_count > 0:
                        events_ago = [i + 1 for i, sat in enumerate(reversed(history)) if sat]
                        events_text = ", ".join(f"{n} event{'s' if n > 1 else ''} ago" for n in events_ago)
                        print(f"{person1.name} and {person2.name} sat together: {events_text}")


def attendees_from_spreadsheet(excelfile):
    try:
        df = pd.read_excel(excelfile)
        df_attending = df[df.attending == 'Y']
        attendees = []
        for ind, row in df_attending.iterrows():
            head_table = True if row['head table'] == 1.0 else False
            attendees.append(
                Attendee(row['name'], row['gender'], row['seniority'], row['division'], assign_head_table=head_table))
        return attendees
    except FileNotFoundError:
        return []