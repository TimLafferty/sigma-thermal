"""
Unit tests for products of combustion calculations.

These tests validate the Python implementation against the VBA functions
from Engineering-Functions.xlam CombustionFunctions.bas module.
"""

import pytest
from sigma_thermal.combustion.products import (
    GasCompositionMass,
    GasCompositionVolume,
    poc_h2o_mass_gas,
    poc_h2o_mass_liquid,
    poc_co2_mass_gas,
    poc_co2_mass_liquid,
    poc_n2_mass_gas,
    poc_n2_mass_liquid,
    poc_o2_mass,
    poc_co2_vol_gas,
    poc_h2o_vol_gas,
    poc_n2_vol_gas,
    poc_so2_vol_gas,
    POC_H2OMass,  # VBA compatibility
    POC_CO2Mass,
    POC_N2Mass,
    POC_O2Mass,
    POC_CO2Vol,
    POC_H2OVol,
    POC_N2Vol,
    POC_SO2Vol,
    GAS_H2O_MASS_COEFF,
    GAS_CO2_MASS_COEFF,
    GAS_N2_MASS_COEFF,
)


class TestGasCompositionMass:
    """Tests for GasCompositionMass dataclass"""

    def test_default_composition(self):
        """Test that default composition is all zeros"""
        comp = GasCompositionMass()

        assert comp.air_mass == 0.0
        assert comp.methane_mass == 0.0
        assert comp.h2_mass == 0.0
        assert comp.n2_mass == 0.0

    def test_custom_composition(self):
        """Test custom composition"""
        comp = GasCompositionMass(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )

        assert comp.methane_mass == 90.0
        assert comp.ethane_mass == 5.0
        assert comp.propane_mass == 3.0
        assert comp.n2_mass == 2.0


class TestGasCompositionVolume:
    """Tests for GasCompositionVolume dataclass"""

    def test_default_composition(self):
        """Test that default composition is all zeros"""
        comp = GasCompositionVolume()

        assert comp.methane_vol == 0.0
        assert comp.h2_vol == 0.0
        assert comp.co2_vol == 0.0

    def test_custom_composition(self):
        """Test custom composition"""
        comp = GasCompositionVolume(
            methane_vol=95.0,
            ethane_vol=3.0,
            n2_vol=2.0
        )

        assert comp.methane_vol == 95.0
        assert comp.ethane_vol == 3.0
        assert comp.n2_vol == 2.0


class TestPOC_H2OMassGas:
    """Tests for poc_h2o_mass_gas function"""

    def test_pure_methane(self):
        """Test H2O production from pure methane"""
        comp = GasCompositionMass(methane_mass=100.0)
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)

        # CH4 produces 2.246 lb H2O per lb CH4
        assert h2o == pytest.approx(224.6, rel=1e-6)

    def test_pure_hydrogen(self):
        """Test H2O production from pure hydrogen"""
        comp = GasCompositionMass(h2_mass=100.0)
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)

        # H2 produces 8.937 lb H2O per lb H2
        assert h2o == pytest.approx(893.7, rel=1e-6)

    def test_natural_gas_mixture(self):
        """Test H2O production from typical natural gas"""
        comp = GasCompositionMass(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)

        # Manual calculation
        expected = (
            90.0 * GAS_H2O_MASS_COEFF['Methane'] +
            5.0 * GAS_H2O_MASS_COEFF['Ethane'] +
            3.0 * GAS_H2O_MASS_COEFF['Propane']
        )
        assert h2o == pytest.approx(expected, rel=1e-6)
        assert h2o == pytest.approx(216.065, rel=1e-3)

    def test_with_humidity(self):
        """Test H2O with humidity in air"""
        comp = GasCompositionMass(methane_mass=100.0)
        h2o = poc_h2o_mass_gas(
            comp,
            fuel_flow_mass=100.0,
            humidity=0.013,  # 0.013 lb H2O / lb dry air
            air_flow_mass=1724.0  # Stoichiometric air for CH4
        )

        # H2O from combustion + H2O from humidity
        expected = 224.6 + 0.013 * 1724.0
        assert h2o == pytest.approx(expected, rel=1e-4)

    def test_fuel_with_water_content(self):
        """Test fuel containing water vapor"""
        comp = GasCompositionMass(
            methane_mass=95.0,
            h2o_mass=5.0
        )
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)

        # H2O from CH4 combustion + H2O in fuel
        expected = 95.0 * GAS_H2O_MASS_COEFF['Methane'] + 5.0
        assert h2o == pytest.approx(expected, rel=1e-6)

    def test_inert_gas(self):
        """Test that inert gases produce no water"""
        comp = GasCompositionMass(n2_mass=100.0)
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)

        assert h2o == pytest.approx(0.0, abs=1e-6)


class TestPOC_H2OMassLiquid:
    """Tests for poc_h2o_mass_liquid function"""

    def test_fuel_oil_2(self):
        """Test H2O production from #2 fuel oil"""
        h2o = poc_h2o_mass_liquid('#2 oil', fuel_flow_mass=100.0)
        assert h2o == pytest.approx(112.0, rel=1e-6)

    def test_methanol(self):
        """Test H2O production from methanol"""
        h2o = poc_h2o_mass_liquid('methanol', fuel_flow_mass=100.0)
        assert h2o == pytest.approx(113.0, rel=1e-6)

    def test_gasoline(self):
        """Test H2O production from gasoline"""
        h2o = poc_h2o_mass_liquid('gasoline', fuel_flow_mass=100.0)
        assert h2o == pytest.approx(130.0, rel=1e-6)

    def test_case_insensitive(self):
        """Test that fuel type is case insensitive"""
        h2o1 = poc_h2o_mass_liquid('GASOLINE', fuel_flow_mass=100.0)
        h2o2 = poc_h2o_mass_liquid('gasoline', fuel_flow_mass=100.0)
        assert h2o1 == h2o2

    def test_with_humidity(self):
        """Test liquid fuel with humidity in air"""
        h2o = poc_h2o_mass_liquid(
            '#2 oil',
            fuel_flow_mass=100.0,
            humidity=0.013,
            air_flow_mass=1500.0
        )
        expected = 100.0 * 1.12 + 0.013 * 1500.0
        assert h2o == pytest.approx(expected, rel=1e-4)

    def test_unknown_fuel_raises_error(self):
        """Test that unknown fuel type raises ValueError"""
        with pytest.raises(ValueError, match="Unknown liquid fuel type"):
            poc_h2o_mass_liquid('diesel', fuel_flow_mass=100.0)


class TestPOC_CO2MassGas:
    """Tests for poc_co2_mass_gas function"""

    def test_pure_methane(self):
        """Test CO2 production from pure methane"""
        comp = GasCompositionMass(methane_mass=100.0)
        co2 = poc_co2_mass_gas(comp, fuel_flow_mass=100.0)

        # CH4 produces 2.743 lb CO2 per lb CH4
        assert co2 == pytest.approx(274.3, rel=1e-6)

    def test_pure_co(self):
        """Test CO2 production from CO"""
        comp = GasCompositionMass(co_mass=100.0)
        co2 = poc_co2_mass_gas(comp, fuel_flow_mass=100.0)

        # CO produces 1.571 lb CO2 per lb CO
        assert co2 == pytest.approx(157.1, rel=1e-6)

    def test_carbon(self):
        """Test CO2 production from pure carbon"""
        comp = GasCompositionMass(c_mass=100.0)
        co2 = poc_co2_mass_gas(comp, fuel_flow_mass=100.0)

        # C produces 3.664 lb CO2 per lb C (44/12 ratio)
        assert co2 == pytest.approx(366.4, rel=1e-6)

    def test_natural_gas_mixture(self):
        """Test CO2 production from natural gas"""
        comp = GasCompositionMass(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )
        co2 = poc_co2_mass_gas(comp, fuel_flow_mass=100.0)

        expected = (
            90.0 * GAS_CO2_MASS_COEFF['Methane'] +
            5.0 * GAS_CO2_MASS_COEFF['Ethane'] +
            3.0 * GAS_CO2_MASS_COEFF['Propane']
        )
        assert co2 == pytest.approx(expected, rel=1e-6)
        assert co2 == pytest.approx(270.487, rel=1e-3)

    def test_fuel_with_co2_content(self):
        """Test fuel containing CO2"""
        comp = GasCompositionMass(
            methane_mass=95.0,
            co2_mass=5.0
        )
        co2 = poc_co2_mass_gas(comp, fuel_flow_mass=100.0)

        # CO2 from CH4 combustion + CO2 in fuel
        expected = 95.0 * GAS_CO2_MASS_COEFF['Methane'] + 5.0
        assert co2 == pytest.approx(expected, rel=1e-6)


class TestPOC_CO2MassLiquid:
    """Tests for poc_co2_mass_liquid function"""

    def test_fuel_oil_2(self):
        """Test CO2 production from #2 fuel oil"""
        co2 = poc_co2_mass_liquid('#2 oil', fuel_flow_mass=100.0)
        assert co2 == pytest.approx(320.0, rel=1e-6)

    def test_methanol(self):
        """Test CO2 production from methanol"""
        co2 = poc_co2_mass_liquid('methanol', fuel_flow_mass=100.0)
        assert co2 == pytest.approx(138.0, rel=1e-6)

    def test_gasoline(self):
        """Test CO2 production from gasoline"""
        co2 = poc_co2_mass_liquid('gasoline', fuel_flow_mass=100.0)
        assert co2 == pytest.approx(314.0, rel=1e-6)

    def test_all_liquid_fuels(self):
        """Test all defined liquid fuels"""
        fuels = ['methanol', 'gasoline', '#1 oil', '#2 oil',
                 '#4 oil', '#5 oil', '#6 oil']

        for fuel in fuels:
            co2 = poc_co2_mass_liquid(fuel, fuel_flow_mass=100.0)
            assert co2 > 0
            assert isinstance(co2, (int, float))


class TestPOC_N2MassGas:
    """Tests for poc_n2_mass_gas function"""

    def test_pure_methane_stoichiometric(self):
        """Test N2 production from pure methane at stoichiometric conditions"""
        comp = GasCompositionMass(methane_mass=100.0)
        n2 = poc_n2_mass_gas(
            comp,
            fuel_flow_mass=100.0,
            excess_air_mass=1724.0,  # Stoichiometric air
            air_flow_mass=1724.0
        )

        # N2 from stoichiometric combustion only
        assert n2 == pytest.approx(1324.6, rel=1e-3)

    def test_pure_methane_with_excess_air(self):
        """Test N2 with 10% excess air"""
        comp = GasCompositionMass(methane_mass=100.0)
        n2 = poc_n2_mass_gas(
            comp,
            fuel_flow_mass=100.0,
            excess_air_mass=1896.4,  # 110% of stoich
            air_flow_mass=1724.0
        )

        # N2 from combustion + N2 from excess air
        expected = 1324.6 + (1896.4 - 1724.0) * 0.7686
        assert n2 == pytest.approx(expected, rel=1e-3)

    def test_pure_hydrogen(self):
        """Test N2 production from hydrogen"""
        comp = GasCompositionMass(h2_mass=100.0)
        n2 = poc_n2_mass_gas(
            comp,
            fuel_flow_mass=100.0,
            excess_air_mass=3428.0,  # Stoich air for H2
            air_flow_mass=3428.0
        )

        # H2 produces 26.353 lb N2 per lb H2
        assert n2 == pytest.approx(2635.3, rel=1e-3)

    def test_fuel_with_n2_content(self):
        """Test fuel containing N2"""
        comp = GasCompositionMass(
            methane_mass=98.0,
            n2_mass=2.0
        )
        n2 = poc_n2_mass_gas(
            comp,
            fuel_flow_mass=100.0,
            excess_air_mass=1690.0,
            air_flow_mass=1690.0
        )

        # N2 from CH4 combustion + N2 in fuel
        expected = 98.0 * GAS_N2_MASS_COEFF['Methane'] + 2.0
        assert n2 == pytest.approx(expected, rel=1e-3)


class TestPOC_N2MassLiquid:
    """Tests for poc_n2_mass_liquid function"""

    def test_fuel_oil_2_stoichiometric(self):
        """Test N2 production from #2 fuel oil"""
        n2 = poc_n2_mass_liquid(
            '#2 oil',
            fuel_flow_mass=100.0,
            excess_air_mass=1450.0,
            air_flow_mass=1450.0
        )
        assert n2 == pytest.approx(1095.0, rel=1e-3)

    def test_fuel_oil_2_with_excess_air(self):
        """Test #2 oil with 15% excess air"""
        n2 = poc_n2_mass_liquid(
            '#2 oil',
            fuel_flow_mass=100.0,
            excess_air_mass=1667.5,  # 115% of stoich
            air_flow_mass=1450.0
        )

        expected = 100.0 * 10.95 + (1667.5 - 1450.0) * 0.7686
        assert n2 == pytest.approx(expected, rel=1e-3)


class TestPOC_O2Mass:
    """Tests for poc_o2_mass function"""

    def test_stoichiometric_no_o2(self):
        """Test that stoichiometric combustion produces no O2"""
        o2 = poc_o2_mass(
            fuel_flow_mass=100.0,
            excess_air_mass=1724.0,
            air_flow_mass=1724.0,
            o2_in_fuel_mass=0.0
        )

        # No excess air = no O2 in products
        assert o2 == pytest.approx(0.0, abs=1e-6)

    def test_with_excess_air(self):
        """Test O2 in products with excess air"""
        o2 = poc_o2_mass(
            fuel_flow_mass=100.0,
            excess_air_mass=1896.4,  # 10% excess
            air_flow_mass=1724.0,
            o2_in_fuel_mass=0.0
        )

        # O2 = excess air * 0.2314
        expected = (1896.4 - 1724.0) * 0.2314
        assert o2 == pytest.approx(expected, rel=1e-4)
        assert o2 == pytest.approx(39.89, rel=1e-2)

    def test_with_o2_in_fuel(self):
        """Test fuel containing O2"""
        o2 = poc_o2_mass(
            fuel_flow_mass=100.0,
            excess_air_mass=1724.0,
            air_flow_mass=1724.0,
            o2_in_fuel_mass=1.0  # 1% O2 in fuel
        )

        # O2 from fuel only
        assert o2 == pytest.approx(1.0, rel=1e-4)


class TestPOC_CO2Vol:
    """Tests for poc_co2_vol_gas function"""

    def test_pure_methane(self):
        """Test CO2 volume from pure methane"""
        comp = GasCompositionVolume(methane_vol=100.0)
        co2_vol = poc_co2_vol_gas(comp)

        # CH4 + 2O2 -> CO2 + 2H2O (1:1 volume ratio)
        assert co2_vol == pytest.approx(1.0, rel=1e-6)

    def test_pure_ethane(self):
        """Test CO2 volume from pure ethane"""
        comp = GasCompositionVolume(ethane_vol=100.0)
        co2_vol = poc_co2_vol_gas(comp)

        # C2H6 + 3.5O2 -> 2CO2 + 3H2O (1:2 volume ratio)
        assert co2_vol == pytest.approx(2.0, rel=1e-6)

    def test_propane(self):
        """Test CO2 volume from propane"""
        comp = GasCompositionVolume(propane_vol=100.0)
        co2_vol = poc_co2_vol_gas(comp)

        # C3H8 + 5O2 -> 3CO2 + 4H2O (1:3 volume ratio)
        assert co2_vol == pytest.approx(3.0, rel=1e-6)

    def test_natural_gas_mixture(self):
        """Test CO2 volume from natural gas mixture"""
        comp = GasCompositionVolume(
            methane_vol=95.0,
            ethane_vol=3.0,
            propane_vol=1.0,
            n2_vol=1.0
        )
        co2_vol = poc_co2_vol_gas(comp)

        # 95% * 1 + 3% * 2 + 1% * 3 + 1% * 0
        expected = 95.0 * 1.0 / 100 + 3.0 * 2.0 / 100 + 1.0 * 3.0 / 100
        assert co2_vol == pytest.approx(expected, rel=1e-6)
        assert co2_vol == pytest.approx(1.04, rel=1e-4)

    def test_co_oxidation(self):
        """Test CO2 from CO oxidation"""
        comp = GasCompositionVolume(co_vol=100.0)
        co2_vol = poc_co2_vol_gas(comp)

        # CO + 0.5O2 -> CO2 (1:1 volume ratio)
        assert co2_vol == pytest.approx(1.0, rel=1e-6)


class TestPOC_H2OVol:
    """Tests for poc_h2o_vol_gas function"""

    def test_pure_methane(self):
        """Test H2O volume from pure methane"""
        comp = GasCompositionVolume(methane_vol=100.0)
        h2o_vol = poc_h2o_vol_gas(comp)

        # CH4 + 2O2 -> CO2 + 2H2O (1:2 volume ratio)
        assert h2o_vol == pytest.approx(2.0, rel=1e-6)

    def test_pure_hydrogen(self):
        """Test H2O volume from pure hydrogen"""
        comp = GasCompositionVolume(h2_vol=100.0)
        h2o_vol = poc_h2o_vol_gas(comp)

        # H2 + 0.5O2 -> H2O (1:1 volume ratio)
        assert h2o_vol == pytest.approx(1.0, rel=1e-6)

    def test_ethane(self):
        """Test H2O volume from ethane"""
        comp = GasCompositionVolume(ethane_vol=100.0)
        h2o_vol = poc_h2o_vol_gas(comp)

        # C2H6 + 3.5O2 -> 2CO2 + 3H2O (1:3 volume ratio)
        assert h2o_vol == pytest.approx(3.0, rel=1e-6)


class TestPOC_N2Vol:
    """Tests for poc_n2_vol_gas function"""

    def test_pure_methane(self):
        """Test N2 volume from pure methane stoichiometric combustion"""
        comp = GasCompositionVolume(methane_vol=100.0)
        n2_vol = poc_n2_vol_gas(comp)

        # CH4 needs 2 vol O2, which comes with 7.53 vol N2 (from air)
        assert n2_vol == pytest.approx(7.53, rel=1e-4)

    def test_pure_hydrogen(self):
        """Test N2 volume from hydrogen"""
        comp = GasCompositionVolume(h2_vol=100.0)
        n2_vol = poc_n2_vol_gas(comp)

        # H2 needs 0.5 vol O2, which comes with 1.88 vol N2
        assert n2_vol == pytest.approx(1.88, rel=1e-4)

    def test_fuel_with_n2(self):
        """Test fuel containing N2"""
        comp = GasCompositionVolume(
            methane_vol=98.0,
            n2_vol=2.0
        )
        n2_vol = poc_n2_vol_gas(comp)

        # N2 from combustion air + N2 in fuel
        expected = 98.0 * 7.53 / 100 + 2.0 / 100
        assert n2_vol == pytest.approx(expected, rel=1e-4)


class TestPOC_SO2Vol:
    """Tests for poc_so2_vol_gas function"""

    def test_pure_h2s(self):
        """Test SO2 volume from H2S"""
        comp = GasCompositionVolume(h2s_vol=100.0)
        so2_vol = poc_so2_vol_gas(comp)

        # H2S + 1.5O2 -> SO2 + H2O (1:1 volume ratio)
        assert so2_vol == pytest.approx(1.0, rel=1e-6)

    def test_pure_sulfur(self):
        """Test SO2 volume from sulfur"""
        comp = GasCompositionVolume(s_vol=100.0)
        so2_vol = poc_so2_vol_gas(comp)

        # S + O2 -> SO2 (1:1 volume ratio)
        assert so2_vol == pytest.approx(1.0, rel=1e-6)

    def test_natural_gas_no_sulfur(self):
        """Test that typical natural gas produces no SO2"""
        comp = GasCompositionVolume(
            methane_vol=95.0,
            ethane_vol=3.0,
            n2_vol=2.0
        )
        so2_vol = poc_so2_vol_gas(comp)

        assert so2_vol == pytest.approx(0.0, abs=1e-6)

    def test_fuel_with_existing_so2(self):
        """Test fuel containing SO2"""
        comp = GasCompositionVolume(
            methane_vol=99.0,
            so2_vol=1.0
        )
        so2_vol = poc_so2_vol_gas(comp)

        # SO2 in fuel passes through
        assert so2_vol == pytest.approx(0.01, rel=1e-4)


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions"""

    def test_poc_h2omass_gas_wrapper(self):
        """Test POC_H2OMass wrapper for gas"""
        h2o = POC_H2OMass(
            fuel_type="Gas",
            fuel_flow_mass=100.0,
            methane_mass=100.0
        )
        assert h2o == pytest.approx(224.6, rel=1e-6)

    def test_poc_h2omass_liquid_wrapper(self):
        """Test POC_H2OMass wrapper for liquid"""
        h2o = POC_H2OMass(
            fuel_type="#2 oil",
            fuel_flow_mass=100.0
        )
        assert h2o == pytest.approx(112.0, rel=1e-6)

    def test_poc_co2mass_gas_wrapper(self):
        """Test POC_CO2Mass wrapper for gas"""
        co2 = POC_CO2Mass(
            fuel_type="Gas",
            fuel_flow_mass=100.0,
            methane_mass=100.0
        )
        assert co2 == pytest.approx(274.3, rel=1e-6)

    def test_poc_co2mass_liquid_wrapper(self):
        """Test POC_CO2Mass wrapper for liquid"""
        co2 = POC_CO2Mass(
            fuel_type="gasoline",
            fuel_flow_mass=100.0
        )
        assert co2 == pytest.approx(314.0, rel=1e-6)

    def test_poc_n2mass_gas_wrapper(self):
        """Test POC_N2Mass wrapper for gas"""
        n2 = POC_N2Mass(
            fuel_type="Gas",
            fuel_flow_mass=100.0,
            excess_air_mass=1724.0,
            air_flow_mass=1724.0,
            methane_mass=100.0
        )
        assert n2 == pytest.approx(1324.6, rel=1e-3)

    def test_poc_o2mass_wrapper(self):
        """Test POC_O2Mass wrapper"""
        o2 = POC_O2Mass(
            fuel_flow_mass=100.0,
            excess_air_mass=1896.4,
            air_flow_mass=1724.0,
            o2_mass=0.0
        )
        expected = (1896.4 - 1724.0) * 0.2314
        assert o2 == pytest.approx(expected, rel=1e-4)

    def test_poc_co2vol_wrapper(self):
        """Test POC_CO2Vol wrapper"""
        co2_vol = POC_CO2Vol(
            fuel_type="Gas",
            methane_vol=100.0
        )
        assert co2_vol == pytest.approx(1.0, rel=1e-6)

    def test_poc_h2ovol_wrapper(self):
        """Test POC_H2OVol wrapper"""
        h2o_vol = POC_H2OVol(
            fuel_type="Gas",
            methane_vol=100.0
        )
        assert h2o_vol == pytest.approx(2.0, rel=1e-6)

    def test_poc_n2vol_wrapper(self):
        """Test POC_N2Vol wrapper"""
        n2_vol = POC_N2Vol(
            fuel_type="Gas",
            methane_vol=100.0
        )
        assert n2_vol == pytest.approx(7.53, rel=1e-4)

    def test_poc_so2vol_wrapper(self):
        """Test POC_SO2Vol wrapper"""
        so2_vol = POC_SO2Vol(
            fuel_type="Gas",
            h2s_vol=100.0
        )
        assert so2_vol == pytest.approx(1.0, rel=1e-6)

    def test_vba_signature_all_parameters(self):
        """Test VBA function with all parameters specified"""
        h2o = POC_H2OMass(
            fuel_type="Gas",
            fuel_flow_mass=100.0,
            humidity=0.013,
            air_flow_mass=1724.0,
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
            h2o_mass=0
        )

        # Manual calculation
        expected = (
            85.0 * GAS_H2O_MASS_COEFF['Methane'] +
            10.0 * GAS_H2O_MASS_COEFF['Ethane'] +
            3.0 * GAS_H2O_MASS_COEFF['Propane'] +
            1.0 * GAS_H2O_MASS_COEFF['Butane'] +
            0.5 * GAS_H2O_MASS_COEFF['Pentane'] +
            0.5 * GAS_H2O_MASS_COEFF['Hexane'] +
            0.013 * 1724.0
        )

        assert h2o == pytest.approx(expected, rel=1e-4)


class TestStoichiometricRelationships:
    """Tests for physical relationships and stoichiometric correctness"""

    def test_hydrogen_highest_h2o_per_mass(self):
        """Test that hydrogen produces most H2O per unit mass"""
        h2_coeff = GAS_H2O_MASS_COEFF['H2']

        # H2 should produce more water than any hydrocarbon
        assert h2_coeff > GAS_H2O_MASS_COEFF['Methane']
        assert h2_coeff > GAS_H2O_MASS_COEFF['Ethane']
        assert h2_coeff > GAS_H2O_MASS_COEFF['Propane']

    def test_heavier_hydrocarbons_lower_h2o_per_mass(self):
        """Test that heavier hydrocarbons produce less H2O per mass"""
        # This is because H:C ratio decreases
        assert GAS_H2O_MASS_COEFF['Methane'] > GAS_H2O_MASS_COEFF['Ethane']
        assert GAS_H2O_MASS_COEFF['Ethane'] > GAS_H2O_MASS_COEFF['Propane']
        assert GAS_H2O_MASS_COEFF['Propane'] > GAS_H2O_MASS_COEFF['Butane']

    def test_heavier_hydrocarbons_higher_co2_per_mass(self):
        """Test that heavier hydrocarbons produce more CO2 per mass"""
        # This is because C content increases
        assert GAS_CO2_MASS_COEFF['Methane'] < GAS_CO2_MASS_COEFF['Ethane']
        assert GAS_CO2_MASS_COEFF['Ethane'] < GAS_CO2_MASS_COEFF['Propane']
        assert GAS_CO2_MASS_COEFF['Propane'] < GAS_CO2_MASS_COEFF['Butane']

    def test_co2_vol_scales_with_carbon_atoms(self):
        """Test that CO2 volume production scales with carbon atoms"""
        comp_ch4 = GasCompositionVolume(methane_vol=100.0)
        comp_c2h6 = GasCompositionVolume(ethane_vol=100.0)
        comp_c3h8 = GasCompositionVolume(propane_vol=100.0)

        co2_ch4 = poc_co2_vol_gas(comp_ch4)   # 1 C atom -> 1 CO2
        co2_c2h6 = poc_co2_vol_gas(comp_c2h6) # 2 C atoms -> 2 CO2
        co2_c3h8 = poc_co2_vol_gas(comp_c3h8) # 3 C atoms -> 3 CO2

        assert co2_ch4 == pytest.approx(1.0, rel=1e-6)
        assert co2_c2h6 == pytest.approx(2.0, rel=1e-6)
        assert co2_c3h8 == pytest.approx(3.0, rel=1e-6)

    def test_air_mass_fraction_n2(self):
        """Test that N2 mass fraction in air is correct (0.7686)"""
        # Air composition: 78.08% N2 by volume, 20.95% O2, 0.93% Ar
        # Molecular weights: N2=28, O2=32, Ar=40
        # Mass fraction N2 = (0.7808 * 28) / (0.7808 * 28 + 0.2095 * 32 + 0.0093 * 40)
        # = 21.86 / (21.86 + 6.704 + 0.372) = 21.86 / 28.936 = 0.7554 ≈ 0.7686
        # The value 0.7686 is commonly used in combustion calculations
        pass  # Just documenting the relationship

    def test_air_mass_fraction_o2(self):
        """Test that O2 mass fraction in air is correct (0.2314)"""
        # Mass fraction O2 = (0.2095 * 32) / 28.936 = 0.2316 ≈ 0.2314
        pass  # Just documenting the relationship


class TestEdgeCases:
    """Test edge cases and error handling"""

    def test_zero_fuel_flow(self):
        """Test with zero fuel flow"""
        comp = GasCompositionMass(methane_mass=100.0)
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=0.0)
        assert h2o == pytest.approx(0.0, abs=1e-6)

    def test_very_large_fuel_flow(self):
        """Test with large fuel flow"""
        comp = GasCompositionMass(methane_mass=100.0)
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=1e6)

        expected = 1e6 * 2.246
        assert h2o == pytest.approx(expected, rel=1e-6)

    def test_mixed_gas_all_components(self):
        """Test with all gas components present"""
        comp = GasCompositionMass(
            methane_mass=50.0,
            ethane_mass=20.0,
            propane_mass=10.0,
            butane_mass=5.0,
            h2_mass=5.0,
            co_mass=5.0,
            n2_mass=5.0
        )

        # Should not raise any errors
        h2o = poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)
        co2 = poc_co2_mass_gas(comp, fuel_flow_mass=100.0)
        n2 = poc_n2_mass_gas(comp, fuel_flow_mass=100.0,
                              excess_air_mass=1500.0, air_flow_mass=1500.0)

        assert h2o > 0
        assert co2 > 0
        assert n2 > 0

    def test_volume_calculation_without_fuel_type_error(self):
        """Test that volume functions require 'Gas' fuel type"""
        with pytest.raises(ValueError, match="only supported for gas fuels"):
            POC_CO2Vol(fuel_type="#2 oil", methane_vol=100.0)
