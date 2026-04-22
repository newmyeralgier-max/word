# -*- coding: utf-8 -*-
"""Post-process V3 (session 7b): добивка склеенного английского.

На входе: текущий GOST-docx (после PR #5).
Делает единственную вещь — массово разделяет слипшиеся английские
слова/словосочетания из словаря Simulink/MATLAB. Работает вне OMML
(формулы не трогаем, кодом в `replace_in_paragraph`).

Основа словаря: все подозрительные токены, которые выдал сканер на
прошлом шаге. Для каждого — вручную подтверждённая раздельная форма.
"""
from __future__ import annotations

import re
import sys

from docx import Document
from docx.oxml.ns import qn

from common import replace_in_paragraph


# Замены. Порядок: более длинные/специфичные — СНАЧАЛА, чтобы не
# сломать подстроки (напр., «Configuration Parameters» раньше, чем
# отдельные «Configuration»).
_REPLACEMENTS: list[tuple[str, str]] = [
    # длинные фразы
    ("LoadFlowandMachineInitializations", "Load Flow and Machine Initializations"),
    ("ImpedancevsFrequencyMeasurements", "Impedance vs Frequency Measurements"),
    ("Magnetizationresistanceandreactance", "Magnetization resistance and reactance"),
    ("Plotsignalsasmagnitudeandphase", "Plot signals as magnitude and phase"),
    ("Specifyimpedanceusingshort-circuitlevel", "Specify impedance using short-circuit level"),
    ("circuitlevelatbasevoltage", "circuit level at base voltage"),
    ("Externalcontrolofswitchingtimes", "External control of switching times"),
    ("Restorecurrentaxessettings", "Restore current axes settings"),
    ("Frequencyofthemodulation", "Frequency of the modulation"),
    ("Initialstatusofbreakers", "Initial status of breakers"),
    ("Initialstatusoffault", "Initial status of fault"),
    ("Showmessagesduringanalysis", "Show messages during analysis"),
    ("Discretizeelectricalmodel", "Discretize electrical model"),
    ("Resistanceperunitlength", "Resistance per unit length"),
    ("Inductanceperunitlength", "Inductance per unit length"),
    ("Plotselectedmeasurements", "Plot selected measurements"),
    ("Nominalpowerandfrequency", "Nominal power and frequency"),
    ("Threewindingstransformer", "Three-windings transformer"),
    ("AvailableMeasurements", "Available Measurements"),
    ("SelectedMeasurements", "Selected Measurements"),
    ("ConfigurationParameters", "Configuration Parameters"),
    ("Showmeasurementport", "Show measurement port"),
    ("Unlockaxesselection", "Unlock axes selection"),
    ("Specifyinitialfluxes", "Specify initial fluxes"),
    ("Restoredisabledlinks", "Restore disabled links"),
    ("Numberofpisections", "Number of pi sections"),
    ("Linesectionlength", "Line section length"),
    ("Internalconnection", "Internal connection"),
    ("Sourceresistance", "Source resistance"),
    ("Snubberresistance", "Snubber resistance"),
    ("Foregroundcolor", "Foreground color"),
    ("Backgroundcolor", "Background color"),
    ("Showdropshadow", "Show drop shadow"),
    ("Showportlabels", "Show port labels"),
    ("Zerocrossingcontrol", "Zero-crossing control"),
    ("Relativetolerance", "Relative tolerance"),
    ("Absolutetolerance", "Absolute tolerance"),
    ("Uselocalsetting", "Use local setting"),
    ("PhaseangleofphaseA", "Phase angle of phase A"),
    ("SwitchingofphaseA", "Switching of phase A"),
    ("Gotoparentsystem", "Go to parent system"),
    ("Textalignment", "Text alignment"),
    ("atsimulationstart", "at simulation start"),
    ("Openatsimulationstart", "Open at simulation start"),
    ("Initialstepsize", "Initial step size"),
    ("Fixedstepsize", "Fixed step size"),
    ("Initialamplitude", "Initial amplitude"),
    ("Initialfrequency", "Initial frequency"),
    ("Initialstatus", "Initial status"),
    ("Switchingtimes", "Switching times"),
    ("Initialfluxes", "Initial fluxes"),
    ("phasermsvoltage", "phase rms voltage"),
    ("Transitiontimes", "Transition times"),
    ("Transitionstatus", "Transition status"),
    ("Externalcontrol", "External control"),
    ("Phasorsimulation", "Phasor simulation"),
    ("Simulationtype", "Simulation type"),
    ("Simulatehysteresis", "Simulate hysteresis"),
    ("Saturablecore", "Saturable core"),
    ("Saturationcharacteristic", "Saturation characteristic"),
    ("Stepmagnitude", "Step magnitude"),
    ("Variationtiming", "Variation timing"),
    ("Harmonictiming", "Harmonic timing"),
    ("Floatingscope", "Floating scope"),
    ("Floatingdisplay", "Floating display"),
    ("Signalselection", "Signal selection"),
    ("Analysistools", "Analysis tools"),
    ("Basevoltage", "Base voltage"),
    ("Initialstate", "Initial state"),
    ("Peakvalues", "Peak values"),
    ("Rateofchange", "Rate of change"),
    ("Opencircuit", "Open circuit"),
    ("Branchtype", "Branch type"),
    ("Sourcetype", "Source type"),
    ("Outputtype", "Output type"),
    ("Outputsignal", "Output signal"),
    ("Signallabel", "Signal label"),
    ("Variablename", "Variable name"),
    ("Screencolor", "Screen color"),
    ("Rotateblock", "Rotate block"),
    ("Minstepsize", "Min step size"),
    ("Phaseshort", "Phase-short"),
    ("Voltagesinp.u.", "Voltages in p.u."),
    ("Currentsinp.u.", "Currents in p.u."),
    ("Voltagesinpu", "Voltages in pu"),
    ("Currentsinpu", "Currents in pu"),
    # c хвостами-символами (Vn, Rm, Lm, Ron, Rg, fn)
    ("MagnetizationresistanceRm", "Magnetization resistance Rm"),
    ("MagnetizationreactanceLm", "Magnetization reactance Lm"),
    ("NominalvoltageVn", "Nominal voltage Vn"),
    ("Nominalfrequencyfn", "Nominal frequency fn"),
    ("BreakerresistanceRon", "Breaker resistance Ron"),
    ("FaultresistanceRon", "Fault resistance Ron"),
    ("GroundresistanceRg", "Ground resistance Rg"),
    # последовательности с capital prefix
    ("sequenceresistances", "sequence resistances"),
    ("sequenceinductances", "sequence inductances"),
    ("sequencecapacitances", "sequence capacitances"),
    # окна настроек, склеенные по-русски тоже (подчистим здесь, раз уж есть)
    ("MaximizeAxes", "Maximize axes"),
    ("Maximizeaxes", "Maximize axes"),
    ("Axesscaling", "Axes scaling"),
    ("AxesScaling", "Axes scaling"),
    ("Time displayoffset", "Time display offset"),
    ("Time-axislabels", "Time-axis labels"),
    ("Sampletime", "Sample time"),
    ("Inputprocessing", "Input processing"),
    ("Time spanoverrunaction", "Time span overrun action"),
    ("Limitdatapointstolast", "Limit data points to last"),
    ("Logdatatoworkspace", "Log data to workspace"),
    ("Display thefullpath", "Display the full path"),
    ("Number ofinputports", "Number of input ports"),
    ("FFTAnalysis", "FFT Analysis"),
    ("Maxstepsize", "Max step size"),
    ("Stoptime", "Stop time"),
    ("Starttime", "Start time"),
]


def apply(body) -> int:
    total = 0
    for p in body.iter(qn("w:p")):
        for old, new in _REPLACEMENTS:
            total += replace_in_paragraph(p, old, new)
    return total


def post_process(in_path: str, out_path: str) -> dict:
    doc = Document(in_path)
    body = doc.element.body
    n = apply(body)
    doc.save(out_path)
    return {"replacements": n}


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: post_process3.py IN.docx OUT.docx")
        sys.exit(1)
    r = post_process(sys.argv[1], sys.argv[2])
    print(r)
