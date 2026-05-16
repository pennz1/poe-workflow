"""
Unit tests for the tier-based machine selection algorithm in app.py.
Tests: prefix_csv_server_names, _get_template_machine_names, snap_budget_to_tier,
       _safe_csv_prefix, learn_tier_machine_selections, get_machine_ids_for_tier,
       load_tier_cache/save_tier_cache, _strip_account_prefix
"""

import sys
import os
import json
import datetime
import tempfile
import unittest
from unittest.mock import patch

# Make sure the app directory is on the path
APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

# We need to mock streamlit and other heavy dependencies before importing app functions
import unittest.mock

# Create mock modules for imports that are not needed for unit testing
mock_st = unittest.mock.MagicMock()
mock_st.secrets = {}
mock_st.set_page_config = unittest.mock.MagicMock()
mock_st.cache_data = lambda f=None, **kwargs: f if f else (lambda fn: fn)

sys.modules['streamlit'] = mock_st
sys.modules['msal'] = unittest.mock.MagicMock()
sys.modules['openai'] = unittest.mock.MagicMock()
sys.modules['docx'] = unittest.mock.MagicMock()
sys.modules['docx.shared'] = unittest.mock.MagicMock()
sys.modules['docx.enum'] = unittest.mock.MagicMock()
sys.modules['docx.enum.text'] = unittest.mock.MagicMock()
sys.modules['docx.oxml'] = unittest.mock.MagicMock()
sys.modules['docx.oxml.ns'] = unittest.mock.MagicMock()
sys.modules['requests'] = unittest.mock.MagicMock()

# Mock frontend.ui module
mock_ui = unittest.mock.MagicMock()
sys.modules['frontend'] = unittest.mock.MagicMock()
sys.modules['frontend.ui'] = mock_ui


# Now we can import the functions we need from app.py
# We import at module level using importlib to handle the st.set_page_config call
import importlib.util

spec = importlib.util.spec_from_file_location("app", os.path.join(APP_DIR, "app.py"))
app_module = importlib.util.module_from_spec(spec)

# Patch st.set_page_config before exec
with patch.object(mock_st, 'set_page_config'):
    try:
        spec.loader.exec_module(app_module)
    except Exception as e:
        # If there are runtime errors from streamlit calls, we can still access functions
        print(f"Note: Module load produced non-fatal error: {type(e).__name__}: {e}")

# Extract the functions we want to test
prefix_csv_server_names = app_module.prefix_csv_server_names
_get_template_machine_names = app_module._get_template_machine_names
snap_budget_to_tier = app_module.snap_budget_to_tier
_safe_csv_prefix = app_module._safe_csv_prefix
learn_tier_machine_selections = app_module.learn_tier_machine_selections
get_machine_ids_for_tier = app_module.get_machine_ids_for_tier
load_tier_cache = app_module.load_tier_cache
save_tier_cache = app_module.save_tier_cache
_strip_account_prefix = app_module._strip_account_prefix
_csv_template_hash = app_module._csv_template_hash
BUDGET_TIERS = app_module.BUDGET_TIERS
TIER_CACHE_PATH = app_module.TIER_CACHE_PATH
_format_usd = app_module._format_usd


class TestPrefixCsvServerNames(unittest.TestCase):
    """Test prefix_csv_server_names()"""

    def test_basic_prefix(self):
        csv_text = (
            "*Server name,IP addresses,*Cores,*Memory (In MB),*OS name\n"
            "VM1,,4,4096,Ubuntu\n"
            "VM2,,8,8192,Windows\n"
            "VM3,,2,2048,Linux\n"
        )
        result = prefix_csv_server_names(csv_text, "contoso")
        lines = result.strip().split("\n")

        # Header should be unchanged
        self.assertEqual(lines[0], "*Server name,IP addresses,*Cores,*Memory (In MB),*OS name")

        # Each server name should get the prefix prepended and contain VM with randomized number
        self.assertTrue(lines[1].startswith("contoso-VM"))
        self.assertTrue(lines[2].startswith("contoso-VM"))
        self.assertTrue(lines[3].startswith("contoso-VM"))

        # Numbers should be randomized (not 1,2,3) but deterministic
        name1 = lines[1].split(",")[0]
        name2 = lines[2].split(",")[0]
        name3 = lines[3].split(",")[0]
        # All names should be different
        self.assertEqual(len({name1, name2, name3}), 3)

        # Same prefix produces same result (deterministic)
        result2 = prefix_csv_server_names(csv_text, "contoso")
        self.assertEqual(result, result2)

        # Different prefix produces different numbers
        result3 = prefix_csv_server_names(csv_text, "fabrikam")
        lines3 = result3.strip().split("\n")
        self.assertNotEqual(lines[1].split(",")[0], lines3[1].split(",")[0])

    def test_rest_of_csv_unchanged(self):
        csv_text = (
            "Header1,Header2,Header3\n"
            "ServerA,data1,data2\n"
            "ServerB,data3,data4\n"
        )
        result = prefix_csv_server_names(csv_text, "myprefix")
        lines = result.strip().split("\n")

        # The rest of columns should remain
        self.assertIn(",data1,data2", lines[1])
        self.assertIn(",data3,data4", lines[2])

    def test_empty_input(self):
        result = prefix_csv_server_names("", "prefix")
        # Should return original (empty string or minimal)
        self.assertIsNotNone(result)

    def test_header_only(self):
        csv_text = "Server name,Cores,Memory\n"
        result = prefix_csv_server_names(csv_text, "test")
        lines = result.strip().split("\n")
        self.assertEqual(lines[0], "Server name,Cores,Memory")


class TestGetTemplateMachineNames(unittest.TestCase):
    """Test _get_template_machine_names()"""

    def test_returns_expected_count(self):
        names = _get_template_machine_names()
        self.assertEqual(len(names), 116, f"Expected 116 machines, got {len(names)}")

    def test_first_machine(self):
        names = _get_template_machine_names()
        self.assertEqual(names[0], "pro-VM1")

    def test_last_machine(self):
        names = _get_template_machine_names()
        self.assertEqual(names[-1], "pro-VM116")

    def test_all_start_with_pro(self):
        names = _get_template_machine_names()
        for name in names:
            self.assertTrue(name.startswith("pro-"), f"Machine name '{name}' doesn't start with 'pro-'")


class TestSnapBudgetToTier(unittest.TestCase):
    """Test snap_budget_to_tier()"""

    def test_low_budget_maps_to_15k(self):
        self.assertEqual(snap_budget_to_tier(10000), 15000)

    def test_exact_15k(self):
        self.assertEqual(snap_budget_to_tier(15000), 15000)

    def test_within_15_percent_of_15k(self):
        # 16000 is within 15% of 15000 (15000 * 1.15 = 17250)
        self.assertEqual(snap_budget_to_tier(16000), 15000)

    def test_beyond_15k_range_maps_to_50k(self):
        # 20000 > 15000*1.15 = 17250, so should go to next tier check
        # 20000 <= 50000*1.15 = 57500, so maps to 50000
        self.assertEqual(snap_budget_to_tier(20000), 50000)

    def test_exact_50k(self):
        self.assertEqual(snap_budget_to_tier(50000), 50000)

    def test_within_15_percent_of_50k(self):
        # 55000 <= 50000*1.15 = 57500
        self.assertEqual(snap_budget_to_tier(55000), 50000)

    def test_beyond_50k_range_maps_to_100k(self):
        # 60000 > 50000*1.15 = 57500 but <= 100000*1.15 = 115000
        self.assertEqual(snap_budget_to_tier(60000), 100000)

    def test_exact_100k(self):
        self.assertEqual(snap_budget_to_tier(100000), 100000)

    def test_200k_maps_to_250k(self):
        # 200000 > 100000*1.15 = 115000 but <= 250000*1.15 = 287500
        self.assertEqual(snap_budget_to_tier(200000), 250000)

    def test_exact_250k(self):
        self.assertEqual(snap_budget_to_tier(250000), 250000)

    def test_300k_exceeds_all_returns_max(self):
        # 300000 > 250000*1.15 = 287500 -> returns last tier
        self.assertEqual(snap_budget_to_tier(300000), 250000)

    def test_none_returns_default(self):
        self.assertEqual(snap_budget_to_tier(None), 250000)

    def test_zero_returns_default(self):
        self.assertEqual(snap_budget_to_tier(0), 250000)

    def test_negative_returns_default(self):
        self.assertEqual(snap_budget_to_tier(-1), 250000)


class TestSafeCsvPrefix(unittest.TestCase):
    """Test _safe_csv_prefix()"""

    def test_simple_name(self):
        self.assertEqual(_safe_csv_prefix("Contoso"), "Contoso")

    def test_spaces_replaced(self):
        result = _safe_csv_prefix("My Company")
        self.assertEqual(result, "My-Company")

    def test_empty_string(self):
        self.assertEqual(_safe_csv_prefix(""), "customer")

    def test_special_chars(self):
        result = _safe_csv_prefix("ABC@#$DEF")
        # Special chars should be replaced with hyphens
        self.assertNotIn("@", result)
        self.assertNotIn("#", result)
        self.assertNotIn("$", result)
        self.assertIn("ABC", result)
        self.assertIn("DEF", result)

    def test_none_input(self):
        self.assertEqual(_safe_csv_prefix(None), "customer")


class TestLearnTierMachineSelections(unittest.TestCase):
    """Test learn_tier_machine_selections()"""

    def _make_mock_machines(self, count=10, prefix="test"):
        """Create mock assessed machines with known costs."""
        machines = []
        for i in range(1, count + 1):
            machines.append({
                "id": f"/subscriptions/sub/machines/{prefix}-pro-VM{i}",
                "properties": {
                    "displayName": f"{prefix}-pro-VM{i}",
                    "monthlyComputeCostForRecommendedSize": 100.0 * i,  # 100, 200, ..., 1000
                    "monthlyStorageCost": 10.0,
                    "monthlyBandwidthCost": 5.0,
                },
            })
        return machines

    def test_returns_dict_with_tiers_key(self):
        machines = self._make_mock_machines()
        progress_msgs = []
        result = learn_tier_machine_selections(machines, "test", lambda msg: progress_msgs.append(msg))
        self.assertIn("tiers", result)
        self.assertIsInstance(result["tiers"], dict)

    def test_all_budget_tiers_present(self):
        machines = self._make_mock_machines()
        result = learn_tier_machine_selections(machines, "test", lambda msg: None)
        for tier in BUDGET_TIERS:
            self.assertIn(str(tier), result["tiers"], f"Tier {tier} not found in result")

    def test_tier_has_required_fields(self):
        machines = self._make_mock_machines()
        result = learn_tier_machine_selections(machines, "test", lambda msg: None)
        for tier_key, tier_data in result["tiers"].items():
            self.assertIn("machine_names", tier_data)
            self.assertIn("machine_count", tier_data)
            self.assertIn("expected_monthly", tier_data)
            self.assertIn("expected_annual", tier_data)

    def test_250k_tier_selects_all_for_small_total(self):
        """With 10 machines at total ~$6900/year, all should be selected for 250K tier."""
        machines = self._make_mock_machines()
        result = learn_tier_machine_selections(machines, "test", lambda msg: None)
        tier_250k = result["tiers"]["250000"]
        # Total monthly: sum(100*i + 10 + 5 for i in 1..10) = sum(115, 215, ..., 1015) = 5650
        # Total annual: 5650 * 12 = 67800
        # 67800 < 250000 * 1.2 = 300000, so all machines selected
        self.assertEqual(tier_250k["machine_count"], 10)

    def test_smaller_tiers_select_fewer_machines(self):
        """For a large total cost, smaller tiers should select fewer machines."""
        # Create expensive machines
        machines = []
        for i in range(1, 11):
            machines.append({
                "id": f"/subscriptions/sub/machines/test-pro-VM{i}",
                "properties": {
                    "displayName": f"test-pro-VM{i}",
                    "monthlyComputeCostForRecommendedSize": 2000.0 * i,
                    "monthlyStorageCost": 100.0,
                    "monthlyBandwidthCost": 50.0,
                },
            })
        result = learn_tier_machine_selections(machines, "test", lambda msg: None)
        # Total monthly: sum(2000*i + 150 for i in 1..10) = 2150+4150+...+20150 = 112500/mo = 1,350,000/yr
        # Smaller tiers should have fewer machines
        tier_15k = result["tiers"]["15000"]
        tier_250k = result["tiers"]["250000"]
        self.assertLess(tier_15k["machine_count"], tier_250k["machine_count"])

    def test_expected_annual_within_budget(self):
        """For each tier, expected_annual should not wildly exceed the tier."""
        machines = self._make_mock_machines(count=10)
        result = learn_tier_machine_selections(machines, "test", lambda msg: None)
        for tier_val in BUDGET_TIERS:
            tier_data = result["tiers"][str(tier_val)]
            # When total cost is small, all machines are selected and cost may be below tier
            # When subset is selected, it should be <= tier * 1.2
            self.assertLessEqual(
                tier_data["expected_annual"],
                tier_val * 1.2 + 1,  # +1 for float rounding
                f"Tier {tier_val}: expected_annual {tier_data['expected_annual']} exceeds tier*1.2"
            )


class TestGetMachineIdsForTier(unittest.TestCase):
    """Test get_machine_ids_for_tier()"""

    def test_returns_correct_ids(self):
        machines = [
            {"id": "id1", "properties": {"displayName": "test-pro-VM1"}},
            {"id": "id2", "properties": {"displayName": "test-pro-VM2"}},
            {"id": "id3", "properties": {"displayName": "test-pro-VM3"}},
        ]
        cache = {
            "tiers": {
                "50000": {
                    "machine_names": ["pro-VM1", "pro-VM3"],
                    "machine_count": 2,
                    "expected_monthly": 500.0,
                    "expected_annual": 6000.0,
                }
            }
        }
        result = get_machine_ids_for_tier(50000, machines, "test", cache)
        self.assertEqual(sorted(result), ["id1", "id3"])

    def test_fallback_when_no_tier_data(self):
        machines = [
            {"id": "id1", "properties": {"displayName": "test-pro-VM1"}},
            {"id": "id2", "properties": {"displayName": "test-pro-VM2"}},
        ]
        cache = {"tiers": {}}
        result = get_machine_ids_for_tier(50000, machines, "test", cache)
        # Should return all machine IDs as fallback
        self.assertEqual(sorted(result), ["id1", "id2"])

    def test_fallback_when_cache_empty(self):
        machines = [
            {"id": "id1", "properties": {"displayName": "x-VM1"}},
        ]
        result = get_machine_ids_for_tier(15000, machines, "x", {})
        self.assertEqual(result, ["id1"])

    def test_fallback_count_when_names_mismatch(self):
        """When cached names don't match, should fall back to selecting by count."""
        machines = [
            {"id": "id1", "properties": {"displayName": "test-pro-VM501"}},
            {"id": "id2", "properties": {"displayName": "test-pro-VM502"}},
            {"id": "id3", "properties": {"displayName": "test-pro-VM503"}},
        ]
        cache = {
            "tiers": {
                "15000": {
                    "machine_names": ["pro-VM1", "pro-VM2"],  # old names, won't match
                    "machine_count": 2,
                    "expected_monthly": 500.0,
                    "expected_annual": 6000.0,
                }
            }
        }
        result = get_machine_ids_for_tier(15000, machines, "test", cache)
        self.assertEqual(len(result), 2)
        self.assertEqual(result, ["id1", "id2"])


class TestLoadSaveTierCache(unittest.TestCase):
    """Test load_tier_cache() / save_tier_cache()"""

    def setUp(self):
        # Use a temp file for cache
        self.original_cache_path = app_module.TIER_CACHE_PATH
        self.temp_cache = tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False)
        self.temp_cache.close()
        app_module.TIER_CACHE_PATH = self.temp_cache.name

    def tearDown(self):
        app_module.TIER_CACHE_PATH = self.original_cache_path
        try:
            os.unlink(self.temp_cache.name)
        except OSError:
            pass

    def test_save_and_load(self):
        test_data = {
            "total_monthly": 5000.0,
            "total_annual": 60000.0,
            "machine_count": 10,
            "tiers": {
                "15000": {"machine_names": ["VM1", "VM2"], "machine_count": 2},
            },
        }
        save_tier_cache(test_data)
        loaded = load_tier_cache()
        self.assertEqual(loaded["total_monthly"], 5000.0)
        self.assertEqual(loaded["tiers"]["15000"]["machine_count"], 2)

    def test_expired_cache_returns_empty(self):
        test_data = {
            "total_monthly": 5000.0,
            "tiers": {},
        }
        save_tier_cache(test_data)

        # Manually modify the created_at to be 8 days ago
        with open(self.temp_cache.name, "r", encoding="utf-8") as f:
            cache = json.load(f)
        old_date = (datetime.datetime.now() - datetime.timedelta(days=8)).isoformat()
        cache["created_at"] = old_date
        with open(self.temp_cache.name, "w", encoding="utf-8") as f:
            json.dump(cache, f)

        loaded = load_tier_cache()
        self.assertEqual(loaded, {})

    def test_wrong_template_hash_returns_empty(self):
        test_data = {
            "total_monthly": 5000.0,
            "tiers": {},
        }
        save_tier_cache(test_data)

        # Modify the template hash to something wrong
        with open(self.temp_cache.name, "r", encoding="utf-8") as f:
            cache = json.load(f)
        cache["template_hash"] = "wrong_hash_12"
        with open(self.temp_cache.name, "w", encoding="utf-8") as f:
            json.dump(cache, f)

        loaded = load_tier_cache()
        self.assertEqual(loaded, {})

    def test_nonexistent_file_returns_empty(self):
        # Point to a file that doesn't exist
        app_module.TIER_CACHE_PATH = "/tmp/nonexistent_test_cache_xyz.json"
        loaded = load_tier_cache()
        self.assertEqual(loaded, {})


class TestStripAccountPrefix(unittest.TestCase):
    """Test _strip_account_prefix()"""

    def test_strips_matching_prefix(self):
        result = _strip_account_prefix("contoso-pro-VM1", "contoso")
        self.assertEqual(result, "pro-VM1")

    def test_no_match_returns_original(self):
        result = _strip_account_prefix("pro-VM1", "contoso")
        self.assertEqual(result, "pro-VM1")

    def test_case_insensitive(self):
        result = _strip_account_prefix("Contoso-pro-VM1", "contoso")
        self.assertEqual(result, "pro-VM1")

    def test_empty_prefix(self):
        result = _strip_account_prefix("pro-VM1", "")
        self.assertEqual(result, "pro-VM1")


class TestCSVTemplateValidation(unittest.TestCase):
    """Validate the Azurecsvtemplate.csv file"""

    def test_header_starts_with_server_name(self):
        csv_path = os.path.join(APP_DIR, "Azurecsvtemplate.csv")
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            header = f.readline().strip()
        first_col = header.split(",")[0]
        self.assertEqual(first_col, "*Server name")

    def test_has_116_data_rows(self):
        csv_path = os.path.join(APP_DIR, "Azurecsvtemplate.csv")
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
        # First line is header, rest are data
        data_rows = lines[1:]
        self.assertEqual(len(data_rows), 116, f"Expected 116 data rows, got {len(data_rows)}")

    def test_all_server_names_start_with_pro(self):
        csv_path = os.path.join(APP_DIR, "Azurecsvtemplate.csv")
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
        for i, line in enumerate(lines[1:], start=2):
            name = line.split(",", 1)[0].strip()
            self.assertTrue(
                name.startswith("pro-"),
                f"Row {i}: server name '{name}' doesn't start with 'pro-'"
            )

    def test_all_have_cores_and_memory(self):
        csv_path = os.path.join(APP_DIR, "Azurecsvtemplate.csv")
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
        # Header: *Server name,IP addresses,*Cores,*Memory (In MB),*OS name,...
        # Columns: 0=Server name, 1=IP, 2=Cores, 3=Memory
        for i, line in enumerate(lines[1:], start=2):
            parts = line.split(",")
            cores = parts[2].strip() if len(parts) > 2 else ""
            memory = parts[3].strip() if len(parts) > 3 else ""
            self.assertTrue(
                cores != "",
                f"Row {i}: Cores field is empty"
            )
            self.assertTrue(
                memory != "",
                f"Row {i}: Memory field is empty"
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
