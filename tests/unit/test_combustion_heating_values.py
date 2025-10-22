"""
Unit tests for heating value calculations.

These tests validate the Python implementation against the VBA functions
from Engineering-Functions.xlam CombustionFunctions.bas module.
"""

import pytest
from sigma_thermal.combustion.heating_values import (
    GasComposition,
    hhv_mass_gas,
    lhv_mass_gas,
    hhv_mass_liquid,
    lhv_mass_liquid,
    HHVMass,  # VBA compatibility
    LHVMass,
    GAS_COMPONENT_HHV,
    GAS_COMPONENT_LHV,
)


class TestGasComposition:
    """Tests for GasComposition dataclass"""

    def test_default_composition(self):
        """Test that default composition is all zeros"""
        comp = GasComposition()
        comp_dict = comp.to_dict()

        assert all(value == 0.0 for value in comp_dict.values())
        assert len(comp_dict) == 16  # 16 components

    def test_custom_composition(self):
        """Test custom composition"""
        comp = GasComposition(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )
        comp_dict = comp.to_dict()

        assert comp_dict['Methane'] == 90.0
        assert comp_dict['Ethane'] == 5.0
        assert comp_dict['Propane'] == 3.0
        assert comp_dict['N2'] == 2.0
        assert sum(comp_dict.values()) == 100.0


class TestHHVMassGas:
    """Tests for hhv_mass_gas function"""

    def test_pure_methane(self):
        """Test HHV for pure methane"""
        comp = GasComposition(methane_mass=100.0)
        hhv = hhv_mass_gas(comp)

        # Should match the HHV of pure methane from lookup table
        assert hhv == pytest.approx(23875.0, rel=1e-6)

    def test_pure_hydrogen(self):
        """Test HHV for pure hydrogen"""
        comp = GasComposition(h2_mass=100.0)
        hhv = hhv_mass_gas(comp)

        assert hhv == pytest.approx(61095.0, rel=1e-6)

    def test_pure_ethane(self):
        """Test HHV for pure ethane"""
        comp = GasComposition(ethane_mass=100.0)
        hhv = hhv_mass_gas(comp)

        assert hhv == pytest.approx(22323.0, rel=1e-6)

    def test_natural_gas_mixture(self):
        """Test HHV for typical natural gas mixture"""
        # Typical composition: 90% CH4, 5% C2H6, 3% C3H8, 2% N2
        comp = GasComposition(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )
        hhv = hhv_mass_gas(comp)

        # Manual calculation
        expected = (90.0 * 23875 + 5.0 * 22323 + 3.0 * 21669 + 2.0 * 0) / 100.0
        assert hhv == pytest.approx(expected, rel=1e-6)
        assert hhv == pytest.approx(23253.72, rel=1e-4)

    def test_inert_gas(self):
        """Test that inert gases contribute zero HHV"""
        comp = GasComposition(n2_mass=50.0, co2_mass=50.0)
        hhv = hhv_mass_gas(comp)

        assert hhv == pytest.approx(0.0, abs=1e-6)

    def test_empty_composition(self):
        """Test empty composition gives zero"""
        comp = GasComposition()
        hhv = hhv_mass_gas(comp)

        assert hhv == pytest.approx(0.0, abs=1e-6)

    def test_all_components(self):
        """Test calculation with all components present"""
        # Equal parts of all combustible components
        total = 0
        comp_dict = {}
        combustibles = ['Methane', 'Ethane', 'Propane', 'H2', 'CO']
        n = len(combustibles)

        for fuel in combustibles:
            comp_dict[fuel.lower() + '_mass'] = 100.0 / n
            total += GAS_COMPONENT_HHV[fuel] / n

        comp = GasComposition(**comp_dict)
        hhv = hhv_mass_gas(comp)

        assert hhv == pytest.approx(total, rel=1e-6)


class TestLHVMassGas:
    """Tests for lhv_mass_gas function"""

    def test_pure_methane(self):
        """Test LHV for pure methane"""
        comp = GasComposition(methane_mass=100.0)
        lhv = lhv_mass_gas(comp)

        assert lhv == pytest.approx(21495.0, rel=1e-6)

    def test_pure_hydrogen(self):
        """Test LHV for pure hydrogen"""
        comp = GasComposition(h2_mass=100.0)
        lhv = lhv_mass_gas(comp)

        # H2 has large difference between HHV and LHV
        assert lhv == pytest.approx(51623.0, rel=1e-6)

    def test_lhv_less_than_hhv(self):
        """Test that LHV is always less than or equal to HHV"""
        comp = GasComposition(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )

        hhv = hhv_mass_gas(comp)
        lhv = lhv_mass_gas(comp)

        assert lhv < hhv
        # Difference should be reasonable (10-15% for hydrocarbons)
        diff_percent = (hhv - lhv) / hhv * 100
        assert 5 < diff_percent < 20

    def test_co_same_hhv_lhv(self):
        """Test that CO has same HHV and LHV (no H2O formed)"""
        comp = GasComposition(co_mass=100.0)

        hhv = hhv_mass_gas(comp)
        lhv = lhv_mass_gas(comp)

        # CO combustion doesn't produce water, so HHV = LHV
        assert hhv == pytest.approx(lhv, rel=1e-6)
        assert hhv == pytest.approx(4347.0, rel=1e-6)


class TestHHVMassLiquid:
    """Tests for hhv_mass_liquid function"""

    def test_fuel_oil_2(self):
        """Test HHV for #2 fuel oil"""
        hhv = hhv_mass_liquid('#2 oil')
        assert hhv == 18993

    def test_gasoline(self):
        """Test HHV for gasoline"""
        hhv = hhv_mass_liquid('gasoline')
        assert hhv == 20190

    def test_methanol(self):
        """Test HHV for methanol"""
        hhv = hhv_mass_liquid('methanol')
        assert hhv == 9797

    def test_case_insensitive(self):
        """Test that fuel type is case insensitive"""
        hhv1 = hhv_mass_liquid('GASOLINE')
        hhv2 = hhv_mass_liquid('gasoline')
        hhv3 = hhv_mass_liquid('Gasoline')

        assert hhv1 == hhv2 == hhv3

    def test_unknown_fuel_raises_error(self):
        """Test that unknown fuel type raises ValueError"""
        with pytest.raises(ValueError, match="Unknown liquid fuel type"):
            hhv_mass_liquid('diesel')

    def test_all_liquid_fuels(self):
        """Test all defined liquid fuels"""
        fuels = ['methanol', 'gasoline', '#1 oil', '#2 oil',
                 '#4 oil', '#5 oil', '#6 oil']

        for fuel in fuels:
            hhv = hhv_mass_liquid(fuel)
            assert hhv > 0
            assert isinstance(hhv, (int, float))


class TestLHVMassLiquid:
    """Tests for lhv_mass_liquid function"""

    def test_fuel_oil_2(self):
        """Test LHV for #2 fuel oil"""
        lhv = lhv_mass_liquid('#2 oil')
        assert lhv == 17855

    def test_methanol(self):
        """Test LHV for methanol"""
        lhv = lhv_mass_liquid('methanol')
        assert lhv == 8706

    def test_lhv_less_than_hhv_all_fuels(self):
        """Test that LHV < HHV for all liquid fuels"""
        fuels = ['methanol', 'gasoline', '#1 oil', '#2 oil',
                 '#4 oil', '#5 oil', '#6 oil']

        for fuel in fuels:
            hhv = hhv_mass_liquid(fuel)
            lhv = lhv_mass_liquid(fuel)
            assert lhv < hhv


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions"""

    def test_hhvmass_gas_wrapper(self):
        """Test HHVMass wrapper for gas"""
        # Pure methane
        hhv = HHVMass(
            fuel_type="Gas",
            methane_mass=100.0
        )
        assert hhv == pytest.approx(23875.0, rel=1e-6)

    def test_hhvmass_liquid_wrapper(self):
        """Test HHVMass wrapper for liquid"""
        hhv = HHVMass(fuel_type="#2 oil")
        assert hhv == 18993

    def test_lhvmass_gas_wrapper(self):
        """Test LHVMass wrapper for gas"""
        # Natural gas mixture
        lhv = LHVMass(
            fuel_type="Gas",
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )

        # Calculate expected
        expected = (90.0 * 21495 + 5.0 * 20418 + 3.0 * 19937 + 2.0 * 0) / 100.0
        assert lhv == pytest.approx(expected, rel=1e-6)

    def test_lhvmass_liquid_wrapper(self):
        """Test LHVMass wrapper for liquid"""
        lhv = LHVMass(fuel_type="gasoline")
        assert lhv == 18790

    def test_vba_signature_all_parameters(self):
        """Test VBA function with all parameters specified"""
        hhv = HHVMass(
            fuel_type="Gas",
            air_mass=0,
            argon_mass=0,
            methane_mass=85.0,
            ethane_mass=10.0,
            propane_mass=3.0,
            butane_mass=1.0,
            pentane_mass=0.5,
            hexane_mass=0.5,
            co2_mass=0,
            co_mass=0,
            c_mass=0,
            n2_mass=0,
            h2_mass=0,
            o2_mass=0,
            h2s_mass=0,
            h2o_mass=0,
        )

        # Calculate expected
        expected = (
            85.0 * 23875 + 10.0 * 22323 + 3.0 * 21669 +
            1.0 * 21321 + 0.5 * 21095 + 0.5 * 20966
        ) / 100.0

        assert hhv == pytest.approx(expected, rel=1e-6)


class TestHeatingValueRelationships:
    """Tests for physical relationships between heating values"""

    def test_hydrogen_has_highest_hhv(self):
        """Test that hydrogen has highest HHV per mass"""
        h2_hhv = GAS_COMPONENT_HHV['H2']

        # Check against all other components
        for component, hhv in GAS_COMPONENT_HHV.items():
            if component != 'H2':
                assert h2_hhv > hhv

    def test_hhv_lhv_difference_proportional_to_hydrogen(self):
        """Test that HHV-LHV difference increases with hydrogen content"""
        # Pure methane (CH4): 25% H by mass
        comp_ch4 = GasComposition(methane_mass=100.0)
        diff_ch4 = hhv_mass_gas(comp_ch4) - lhv_mass_gas(comp_ch4)

        # Pure hydrogen: 100% H by mass
        comp_h2 = GasComposition(h2_mass=100.0)
        diff_h2 = hhv_mass_gas(comp_h2) - lhv_mass_gas(comp_h2)

        # H2 should have larger difference (more H2O formed)
        assert diff_h2 > diff_ch4

    def test_heavier_hydrocarbons_lower_hhv(self):
        """Test that heavier hydrocarbons have lower HHV per mass"""
        # Generally true: lighter hydrocarbons have higher H:C ratio
        assert GAS_COMPONENT_HHV['Methane'] > GAS_COMPONENT_HHV['Ethane']
        assert GAS_COMPONENT_HHV['Ethane'] > GAS_COMPONENT_HHV['Propane']
        assert GAS_COMPONENT_HHV['Propane'] > GAS_COMPONENT_HHV['Butane']

    def test_liquid_fuel_order(self):
        """Test expected order for liquid fuels"""
        # Gasoline should have higher HHV than fuel oils
        gasoline_hhv = hhv_mass_liquid('gasoline')
        oil2_hhv = hhv_mass_liquid('#2 oil')
        oil6_hhv = hhv_mass_liquid('#6 oil')

        assert gasoline_hhv > oil2_hhv
        assert oil2_hhv > oil6_hhv  # Heavier oils have lower HHV


class TestEdgeCases:
    """Test edge cases and error handling"""

    def test_composition_over_100_percent(self):
        """Test that >100% composition still calculates (no validation yet)"""
        # Note: Current implementation doesn't validate sum = 100%
        # This is intentional to match VBA behavior
        comp = GasComposition(methane_mass=150.0)
        hhv = hhv_mass_gas(comp)

        # Should be 150% of methane HHV
        assert hhv == pytest.approx(1.5 * 23875, rel=1e-6)

    def test_negative_composition(self):
        """Test that negative compositions are allowed (match VBA)"""
        # VBA doesn't validate, so we don't either
        comp = GasComposition(methane_mass=-10.0)
        hhv = hhv_mass_gas(comp)

        assert hhv < 0  # Negative composition gives negative HHV

    def test_very_small_values(self):
        """Test with very small composition values"""
        comp = GasComposition(methane_mass=0.001)
        hhv = hhv_mass_gas(comp)

        expected = 23875 * 0.001 / 100.0
        assert hhv == pytest.approx(expected, rel=1e-6)
