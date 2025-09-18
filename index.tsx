/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/
import { GoogleGenAI, Type } from "@google/genai";

// Declare the Gemini AI client instance. It will be initialized upon admin login.
let ai: GoogleGenAI | null = null;

// Fix: Moved helper functions to a higher scope to resolve 'Cannot find name' errors.
const getEl = (id: string) => document.getElementById(id);
const getVal = (id: string) => (getEl(id) as HTMLInputElement)?.value || '';
const getSelectedText = (id: string) => {
    const select = getEl(id) as HTMLSelectElement;
    if (!select || select.selectedIndex < 0) {
        return 'N/A';
    }
    return select.options[select.selectedIndex].text;
};

// Fix: Moved `getRadioVal` and `getRadioLabel` to the global scope to be accessible by all functions.
const getRadioVal = (name: string) => {
    const checked = document.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`);
    return checked ? checked.value : '';
};
const getRadioLabel = (name: string) => {
    const checked = document.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`);
    return checked ? (document.querySelector(`label[for="${checked.id}"]`)?.textContent || 'N/A') : 'N/A';
};

// Moved from inside DOMContentLoaded to be globally available.
const getFormData = () => {
  const connectionsSelect = getEl('connectionTypes') as HTMLSelectElement;
  const selectedOpts = Array.from(connectionsSelect.selectedOptions).map(opt => opt.text);
  let connectionsValue = selectedOpts.filter(o => o !== 'Other (specify)...');
  if (selectedOpts.includes('Other (specify)...')) {
      connectionsValue.push(`Other: ${getVal('customConnectionTypes') || 'not specified'}`);
  }

  let heatSourcePipeMaterial = getSelectedText('pipeMaterial').replace(/ \(.*\)/, '');
  if (getVal('pipeMaterial') === 'other') {
      heatSourcePipeMaterial = `Other: ${getVal('customPipeMaterial')}`;
  }

  let heatSourceWallThickness = getVal('wallThickness');
  if (heatSourceWallThickness === 'custom') {
      heatSourceWallThickness = getVal('customWallThickness');
  }

  let insulationType = 'N/A';
  const insulationVal = getVal('pipeInsulationType');
  if (insulationVal === 'none') {
      insulationType = 'None';
  } else if (insulationVal === 'other') {
      insulationType = `Other: ${getVal('customPipeInsulation')}`;
  } else if (insulationVal) {
      insulationType = getSelectedText('pipeInsulationType').replace(/ \(.*\)/, '');
  }

  let gasOperatorName = getSelectedText('gasOperatorName');
  if (getVal('gasOperatorName') === 'other') {
      gasOperatorName = `Other: ${getVal('customGasOperatorName')}`;
  }

  let gasPipeMaterial = getSelectedText('gasPipeMaterial');
  const gasPipeMaterialValue = getVal('gasPipeMaterial');
  if (gasPipeMaterialValue === 'other') {
      gasPipeMaterial = `Other: ${getVal('customGasPipeMaterial')}`;
  }
  
  let gasCoatingType = 'N/A';
  let gasCoatingMaxTemp = 'N/A';
  let gasPipeContinuousLimit = 'N/A';
  let gasPipeUtilityCap = 'N/A';
  let gasPipeNotes = 'N/A';
  
  const gasMaterialSelect = getEl('gasPipeMaterial') as HTMLSelectElement;
  if (gasPipeMaterialValue === 'coated-steel-protected' || gasPipeMaterialValue === 'coated-steel-unprotected') {
      const coatingSelect = getEl('gasCoatingType') as HTMLSelectElement;
      if (coatingSelect.value === 'other') {
          gasCoatingType = `Other: ${getVal('customGasCoatingType')}`;
          gasCoatingMaxTemp = 'Custom';
      } else if (coatingSelect.selectedIndex >= 0 && coatingSelect.value) {
          const selectedOption = coatingSelect.options[coatingSelect.selectedIndex];
          gasCoatingType = selectedOption.text;
          gasCoatingMaxTemp = selectedOption.dataset.maxTemp || 'N/A';
      }
  } else if (['hdpe', 'mdpe', 'aldyl'].includes(gasPipeMaterialValue)) {
      if (gasMaterialSelect.selectedIndex >= 0) {
        const selectedOption = gasMaterialSelect.options[gasMaterialSelect.selectedIndex];
        gasPipeContinuousLimit = selectedOption.dataset.continuousLimit || 'N/A';
        gasPipeUtilityCap = selectedOption.dataset.utilityCap || 'N/A';
        gasPipeNotes = selectedOption.dataset.notes || 'N/A';
      }
  }
  
  const isHeatSourceApplicable = ['steam', 'hotWater', 'superHeatedHotWater', 'other'].includes(getRadioVal('heatSourceType'));

  let heatLossEvidenceSource = getRadioLabel('heatLossEvidenceSource');
  if (getRadioVal('heatLossEvidenceSource') === 'other') {
      heatLossEvidenceSource = `Other: ${getVal('customHeatLossSource')}`;
  }

  let insulationConditionSource = getRadioLabel('insulationConditionSource');
  if (getVal('insulationConditionSource') === 'other') {
    insulationConditionSource = `Other: ${getVal('customInsulationConditionSource')}`;
  }

  const orientation = getRadioVal('gasLineOrientation');
  const parallelCoordinates: { label: string; lat: string; lng: string }[] = [];
  if (orientation === 'parallel') {
      const latInputs = document.querySelectorAll<HTMLInputElement>('#parallelCoordinatesWrapper input[id^="lat-parallel-"]');
      latInputs.forEach(latInput => {
          const index = latInput.id.split('-').pop();
          const lngInput = document.getElementById(`lng-parallel-${index}`) as HTMLInputElement;
          if (lngInput) {
            parallelCoordinates.push({
                label: latInput.dataset.pointLabel || 'N/A',
                lat: latInput.value,
                lng: lngInput.value
            });
          }
      });
  }

  let heatSourceType = getRadioLabel('heatSourceType');
  if (getRadioVal('heatSourceType') === 'other') {
    heatSourceType = `Other: ${getVal('customHeatSourceType') || 'not specified'}`;
  }

  let gasPipeWallThickness = 'N/A';
  const gasWallThicknessGroup = document.getElementById('gasWallThicknessGroup');
  if (gasWallThicknessGroup && !gasWallThicknessGroup.classList.contains('hidden')) {
      const wallThicknessVal = getVal('gasWallThickness');
      if (wallThicknessVal === 'custom') {
          gasPipeWallThickness = getVal('customGasWallThickness');
      } else if (wallThicknessVal) {
          gasPipeWallThickness = wallThicknessVal;
      }
  }

  let gasPipeSDR = 'N/A';
  const gasSdrGroup = document.getElementById('gasSdrGroup');
  if (gasSdrGroup && !gasSdrGroup.classList.contains('hidden')) {
      const sdrVal = getVal('gasSdr');
      if (sdrVal === 'other') {
          gasPipeSDR = getVal('customGasSdr');
      } else if (sdrVal) {
          gasPipeSDR = sdrVal;
      }
  }

  return {
      // Evaluation
      date: getVal('date'),
      evaluatorNames: Array.from(document.querySelectorAll<HTMLInputElement>('.evaluator-name-input')).map(input => input.value),
      engineerName: getVal('engineerName'),
      projectName: getVal('projectName'),
      projectLocation: getVal('projectLocation'),
      projectDescription: getVal('projectDescription'),
      // Heat Source
      isHeatSourceApplicable,
      heatSourceType,
      operatorName: getVal('operatorName'),
      operatorCompanyName: getVal('operatorCompanyName'),
      operatorCompanyAddress: getVal('operatorCompanyAddress'),
      operatorContactInfo: getVal('operatorContactInfo'),
      isRegistered811: getRadioLabel('isRegistered811'),
      registered811Confirmation: getRadioLabel('registered811Confirmation'),
      confirmationDate: getVal('confirmationDate'),
      tempType: getRadioLabel('tempType'),
      maxTemp: getVal('maxTemp'),
      pressureType: getRadioLabel('pressureType'),
      maxPressure: getVal('maxPressure'),
      ageType: getRadioLabel('ageType'),
      heatSourceAge: getVal('heatSourceAge'),
      systemDutyCycleType: getRadioLabel('dutyCycleType'),
      systemDutyCycle: getVal('systemDutyCycle'),
      pipeCasingInfoType: getRadioLabel('casingInfoType'),
      pipeCasingInfo: getVal('pipeCasingInfo'),
      heatLossEvidenceSource,
      heatLossEvidence: getVal('heatLossEvidence'),
      diameterType: getRadioLabel('diameterType'),
      pipelineDiameter: getVal('pipelineDiameter'),
      materialType: getRadioLabel('materialType'),
      heatSourcePipeMaterial,
      wallThicknessType: getRadioLabel('wallThicknessType'),
      heatSourceWallThickness,
      connectionTypesType: getRadioLabel('connectionTypesType'),
      connectionsValue,
      insulationTypeType: getRadioLabel('insulationTypeType'),
      insulationType,
      heatSourcePipeInsulationType: getVal('pipeInsulationType'), // Use canonical key for saving
      customInsulationThermalConductivity: getVal('customInsulationThermalConductivity'),
      insulationThickness: getVal('insulationThickness'),
      insulationThicknessType: getRadioLabel('insulationThicknessType'),
      insulationConditionSource,
      insulationCondition: getVal('insulationCondition'),
      depthType: getRadioLabel('depthType'),
      heatSourceDepth: getVal('heatSourceDepth'),
      additionalInfo: getVal('additionalInfo'),
      // Gas Line
      gasOperatorName,
      gasMaxPressure: getVal('gasMaxPressure'),
      gasInstallationYear: getVal('gasInstallationYear'),
      gasPipelineDiameter: getVal('gasPipelineDiameter'),
      gasPipeSizingStandard: getRadioVal('gasPipeSizingStandard'),
      gasPipeMaterial,
      gasPipeMaterialValue,
      customGasThermalConductivity: getVal('customGasThermalConductivity'),
      gasPipeWallThickness,
      gasPipeSDR,
      gasCoatingType,
      gasCoatingMaxTemp,
      gasPipeContinuousLimit,
      gasPipeUtilityCap,
      gasPipeNotes,
      gasLineOrientation: getRadioLabel('gasLineOrientation'),
      depthOfBurialGasLine: getVal('depthOfBurialGasLine'),
      parallelDistance: orientation === 'parallel' ? getVal('parallelDistance') : 'N/A',
      parallelLength: orientation === 'parallel' ? getVal('parallelLength') : 'N/A',
      parallelCoordinates,
      latitude: orientation === 'perpendicular' ? getVal('latitude') : 'N/A',
      longitude: orientation === 'perpendicular' ? getVal('longitude') : 'N/A',
      lateralOffset: orientation === 'perpendicular' ? getVal('lateralOffset') : 'N/A',
      crossingAngle: orientation === 'perpendicular' ? getVal('crossingAngle') : 'N/A',
      // Soil
      soilType: getVal('soilType'),
      soilThermalConductivity: getVal('soilThermalConductivity'),
      soilMoistureContent: getVal('soilMoistureContent'),
      averageGroundTemperature: getVal('averageGroundTemperature'),
      waterInfiltration: getRadioLabel('waterInfiltration'),
      waterInfiltrationComments: getVal('waterInfiltrationComments'),
      heatSourceBeddingType: getRadioVal('heatSourceBeddingType'),
      heatSourceBeddingTop: getVal('heatSourceBeddingTop'),
      heatSourceBeddingBottom: getVal('heatSourceBeddingBottom'),
      heatSourceBeddingLeft: getVal('heatSourceBeddingLeft'),
      heatSourceBeddingRight: getVal('heatSourceBeddingRight'),
      heatSourceBeddingUseCustomK: getRadioVal('heatSourceBeddingUseCustomK'),
      heatSourceBeddingCustomK: getVal('heatSourceBeddingCustomK'),
      gasLineBeddingType: getRadioVal('gasLineBeddingType'),
      gasBeddingTop: getVal('gasBeddingTop'),
      gasBeddingBottom: getVal('gasBeddingBottom'),
      gasBeddingLeft: getVal('gasBeddingLeft'),
      gasBeddingRight: getVal('gasBeddingRight'),
      gasBeddingUseCustomK: getRadioVal('gasBeddingUseCustomK'),
      gasBeddingCustomK: getVal('gasBeddingCustomK'),
      // Field Visit
      visitDate: getVal('visitDate'),
      sitePersonnel: getVal('sitePersonnel'),
      siteConditions: getVal('siteConditions'),
      fieldObservations: getVal('fieldObservations'),
      // Recommendations
      recommendations: Array.from(document.querySelectorAll<HTMLTextAreaElement>('#recommendationsContainer .recommendation-textarea')).map(textarea => textarea.value),
  };
};

const escapeLatex = (str: string | number | undefined | null): string => {
    if (str === null || typeof str === 'undefined' || str === '') {
        return 'N/A';
    }
    return String(str)
        .replace(/&/g, '\\&')
        .replace(/%/g, '\\%')
        .replace(/\$/g, '\\$')
        .replace(/#/g, '\\#')
        .replace(/_/g, '\\_')
        .replace(/{/g, '\\{')
        .replace(/}/g, '\\}')
        .replace(/~/g, '\\textasciitilde{}')
        .replace(/\^/g, '\\textasciicircum{}')
        .replace(/\\/g, '\\textbackslash{}')
        .replace(/\n/g, '\\\\ ');
};

const generateLatexReport = (data: ReturnType<typeof getFormData>, asIsResults: any, worstCaseResults: any) => {
    const tableRow = (key: string, value: any) => `${key} & ${escapeLatex(value)} \\\\`;
    const multiLineRow = (key: string, value: any) => `${key} & \\parbox[t]{0.6\\textwidth}{${escapeLatex(value)}} \\\\`;

    const preamble = `\\documentclass[11pt, letterpaper]{article}
\\usepackage[margin=1in]{geometry}
\\usepackage{amsmath}
\\usepackage{graphicx}
\\usepackage{booktabs}
\\usepackage{longtable}
\\usepackage{xcolor}
\\usepackage[hidelinks]{hyperref}
\\usepackage{fancyhdr}
\\usepackage{titlesec}
\\usepackage{enumitem}

\\titlespacing*{\\section}{0pt}{1.5ex plus 0.5ex minus .2ex}{1.5ex plus .2ex}
\\titlespacing*{\\subsection}{0pt}{1.5ex plus 0.5ex minus .2ex}{1.5ex plus .2ex}

\\pagestyle{fancy}
\\fancyhf{}
\\fancyhead[L]{Thermal–Gas Line Encroachment Assessment}
\\fancyfoot[C]{\\thepage}

\\title{Thermal–Gas Line Encroachment Assessment Report}
\\author{${escapeLatex(data.engineerName)} \\\\ Rev-2.0, By Tim Bickford}
\\date{${escapeLatex(data.date)}}
`;

    const beginDoc = `\\begin{document}
\\maketitle
\\thispagestyle{empty}
\\newpage
\\tableofcontents
\\newpage
`;

    const asIsFinalTemp = !isNaN(asIsResults.T_gas_line_layered) ? asIsResults.T_gas_line_layered : asIsResults.T_gas_line;
    let worstCaseFinalTemp = NaN;
    if (worstCaseResults) {
        worstCaseFinalTemp = !isNaN(worstCaseResults.T_gas_line_layered) ? worstCaseResults.T_gas_line_layered : worstCaseResults.T_gas_line;
    }

    const summary = `\\section{Executive Summary}
This report details the thermal analysis of a potential encroachment between a buried heat source (${escapeLatex(data.heatSourceType)}) and a natural gas pipeline. The assessment was conducted using the data provided, based on two-dimensional steady-state heat transfer models for a ${escapeLatex(data.gasLineOrientation)} orientation.

The primary finding for the \\textbf{as-is condition} is a calculated gas line temperature of \\textbf{${asIsFinalTemp.toFixed(1)}~$^{\\circ}$F}, with a ground surface temperature (1 inch below grade, above the heat source) of \\textbf{${asIsResults.T_ground_surface_above_hs.toFixed(1)}~$^{\\circ}$F}.
${worstCaseResults ? `
A worst-case scenario, assuming failure or absence of insulation on the heat source line, was also analyzed. This scenario resulted in a calculated gas line temperature of \\textbf{${worstCaseFinalTemp.toFixed(1)}~$^{\\circ}$F} and a surface temperature of \\textbf{${worstCaseResults.T_ground_surface_above_hs.toFixed(1)}~$^{\\circ}$F}.` : ''}

These results are compared against the operational limits of the gas pipeline material and its coating to determine potential risks. Recommendations based on these findings are provided in Section \\ref{sec:recommendations}.
`;

    const inputData = `\\section{Input Data Summary}
\\subsection{Evaluation Information}
\\begin{longtable}{p{0.3\\textwidth} p{0.6\\textwidth}}
\\toprule
\\textbf{Parameter} & \\textbf{Value} \\\\
\\midrule
\\endhead
\\bottomrule
\\endfoot
${tableRow('Project Name', data.projectName)}
${tableRow('Project Location', data.projectLocation)}
${tableRow('Date of Assessment', data.date)}
${tableRow('Engineer Name', data.engineerName)}
${tableRow('Evaluator(s)', data.evaluatorNames.join(', '))}
${multiLineRow('Project Description', data.projectDescription)}
\\end{longtable}

\\subsection{Heat Source Data}
\\begin{longtable}{p{0.3\\textwidth} p{0.6\\textwidth}}
\\toprule
\\textbf{Parameter} & \\textbf{Value} \\\\
\\midrule
\\endhead
\\bottomrule
\\endfoot
${data.isHeatSourceApplicable ? `
${tableRow('Heat Source Type', data.heatSourceType)}
${tableRow('Operator Company Name', data.operatorCompanyName)}
${tableRow('Operating Company Address', data.operatorCompanyAddress)}
${tableRow('Data Provider Name', data.operatorName)}
${tableRow('Provider Contact Info', data.operatorContactInfo)}
${tableRow('811 "DigSafe" Registration', `${data.isRegistered811} (${data.registered811Confirmation})`)}
${tableRow('Data Confirmation Date', data.confirmationDate)}
${tableRow('Max Operating Temp ($^{\\circ}$F)', `${data.maxTemp} (${data.tempType})`)}
${tableRow('Max Operating Pressure (psig)', `${data.maxPressure} (${data.pressureType})`)}
${tableRow('Line Age (years)', `${data.heatSourceAge} (${data.ageType})`)}
${multiLineRow('System Duty Cycle', `${data.systemDutyCycle} (${data.systemDutyCycleType})`)}
${multiLineRow('Pipe Casing / Conduit Info', `${data.pipeCasingInfo} (${data.pipeCasingInfoType})`)}
${multiLineRow('Evidence of Surface Heat Loss', `${data.heatLossEvidence} (Source: ${data.heatLossEvidenceSource})`)}
${tableRow('Nominal Diameter (in)', `${getSelectedText('pipelineDiameter')} (${data.diameterType})`)}
${tableRow('Pipe Material', `${data.heatSourcePipeMaterial} (${data.materialType})`)}
${tableRow('Pipe Wall Thickness (in)', `${data.heatSourceWallThickness} (${data.wallThicknessType})`)}
${multiLineRow('Line Connection Types', `${data.connectionsValue.join(', ')} (${data.connectionTypesType})`)}
${tableRow('Insulation Type', `${data.insulationType} (${data.insulationTypeType})`)}
${data.insulationType !== 'None' ? tableRow('Insulation Thickness (in)', `${data.insulationThickness} (${data.insulationThicknessType})`) : ''}
${multiLineRow('Known Insulation Condition', `${data.insulationCondition} (Source: ${data.insulationConditionSource})`)}
${tableRow('Depth to Centerline (ft)', `${data.heatSourceDepth} (${data.depthType})`)}
${multiLineRow('Additional Operator Info', data.additionalInfo)}
` : tableRow('Heat Source Status', 'No applicable heat source selected.')}
\\end{longtable}

\\subsection{Gas Line Data}
\\begin{longtable}{p{0.3\\textwidth} p{0.6\\textwidth}}
\\toprule
\\textbf{Parameter} & \\textbf{Value} \\\\
\\midrule
\\endhead
\\bottomrule
\\endfoot
${tableRow('Operator Name', data.gasOperatorName)}
${tableRow('Installation Year', data.gasInstallationYear)}
${tableRow('Max Operating Pressure (psig)', data.gasMaxPressure)}
${tableRow('Nominal Diameter (in)', getSelectedText('gasPipelineDiameter'))}
${tableRow('Pipe Sizing Standard', data.gasPipeSizingStandard.toUpperCase())}
${tableRow('Pipe Material', data.gasPipeMaterial)}
${data.gasPipeWallThickness !== 'N/A' ? tableRow('Pipe Wall Thickness (in)', data.gasPipeWallThickness) : ''}
${data.gasPipeSDR !== 'N/A' ? tableRow('Pipe SDR', data.gasPipeSDR) : ''}
${data.gasCoatingType !== 'N/A' ? tableRow('Coating Type', data.gasCoatingType) : ''}
${data.gasCoatingMaxTemp !== 'N/A' ? tableRow('Coating Max Temp ($^{\\circ}$F)', data.gasCoatingMaxTemp) : ''}
${data.gasPipeContinuousLimit !== 'N/A' ? tableRow('Material Continuous Temp Limit ($^{\\circ}$F)', data.gasPipeContinuousLimit) : ''}
${multiLineRow('Key Material Notes', data.gasPipeNotes)}
${tableRow('Orientation to Heat Source', data.gasLineOrientation)}
${tableRow('Depth to Centerline (ft)', data.depthOfBurialGasLine)}
${data.gasLineOrientation === 'Parallel' ? tableRow('Parallel Centerline Distance (ft)', data.parallelDistance) : ''}
${data.gasLineOrientation === 'Parallel' ? tableRow('Parallel Length (ft)', data.parallelLength) : ''}
${data.gasLineOrientation === 'Crossing / Perpendicular' ? tableRow('Lateral Offset (ft)', data.lateralOffset) : ''}
${data.gasLineOrientation === 'Crossing / Perpendicular' ? tableRow('Crossing Angle (deg)', data.crossingAngle) : ''}
${(data.latitude !== 'N/A' && data.longitude !== 'N/A') ? tableRow('Intersection Coordinates', `${data.latitude}, ${data.longitude}`) : ''}
\\end{longtable}

\\subsection{Soil and Bedding Data}
\\begin{longtable}{p{0.3\\textwidth} p{0.6\\textwidth}}
\\toprule
\\textbf{Parameter} & \\textbf{Value} \\\\
\\midrule
\\endhead
\\bottomrule
\\endfoot
${tableRow('Native Soil Type', getSelectedText('soilType').replace(/ \\(.*\\)/, ''))}
${tableRow('Native Soil Thermal Conductivity (BTU/hr$\\cdot$ft$\\cdot$$^{\\circ}$F)', data.soilThermalConductivity)}
${tableRow('Soil Moisture Content (\\%)', data.soilMoistureContent)}
${tableRow('Average Ground Temp ($^{\\circ}$F)', data.averageGroundTemperature)}
${tableRow('Evidence of Water Infiltration', data.waterInfiltration)}
${multiLineRow('Infiltration Comments', data.waterInfiltrationComments)}
${tableRow('Heat Source Bedding', data.heatSourceBeddingType === 'sand' ? `Sand, k = ${data.heatSourceBeddingUseCustomK === 'yes' ? data.heatSourceBeddingCustomK : '0.20'}` : 'Native Soil')}
${tableRow('Gas Line Bedding', data.gasLineBeddingType === 'sand' ? `Sand, k = ${data.gasBeddingUseCustomK === 'yes' ? data.gasBeddingCustomK : '0.20'}` : 'Native Soil')}
\\end{longtable}

\\subsection{Field Visit Data}
\\begin{longtable}{p{0.3\\textwidth} p{0.6\\textwidth}}
\\toprule
\\textbf{Parameter} & \\textbf{Value} \\\\
\\midrule
\\endhead
\\bottomrule
\\endfoot
${tableRow('Date of Visit', data.visitDate)}
${multiLineRow('Personnel on Site', data.sitePersonnel)}
${multiLineRow('Weather and Site Conditions', data.siteConditions)}
${multiLineRow('Field Observations and Notes', data.fieldObservations)}
\\end{longtable}
`;

    const analysis = `\\section{Heat Transfer Analysis}
\\subsection{Governing Equations}
The analysis is based on a two-dimensional, steady-state heat transfer model using the thermal resistance analogy.

\\subsubsection{Total Heat Loss (Q)}
The rate of heat loss per unit length from the heat source is calculated as the temperature difference divided by the total thermal resistance.
\\begin{equation}
Q = \\frac{T_{hs} - T_{surface}}{R_{total}}
\end{equation}

\\subsubsection{Thermal Resistances (R)}
The total resistance is the sum of component resistances from the pipe interior to the ground surface.
\\begin{align}
R_{total} &= R_{pipe} + R_{ins} + R_{bed,hs} + R_{soil} \\\\
R_{pipe} &= \\frac{\\ln(D_{hs,o} / D_{hs,i})}{2\\pi k_{pipe}} && \\text{(Pipe Wall)} \\\\
R_{ins} &= \\frac{\\ln(D_{ins,o} / D_{hs,o})}{2\\pi k_{ins}} && \\text{(Insulation)} \\\\
R_{bed,hs} &= \\frac{\\ln(D_{bed,o} / D_{ins,o})}{2\\pi k_{bed,hs}} && \\text{(Heat Source Bedding)} \\\\
R_{soil} &= \\frac{\\ln((2Z_{hs} - r_{soil,i}) / r_{soil,i})}{2\\pi k_{soil}} && \\text{(Soil, Method of Images)}
\end{align}

\\subsubsection{Gas Line Temperature ($T_{gas}$)}
The temperature at the gas line is the sum of the ambient ground temperature and the temperature rise caused by the heat source. This base calculation assumes a homogeneous soil environment.

\\textbf{Perpendicular/Crossing Orientation ($T_{perp}$):}
\\begin{equation}
T_{perp} = T_{surface} + \\frac{Q}{2\\pi k_{soil}} \\ln\\left(\\frac{Z_{hs} + Z_{gas}}{D}\\right)
\end{equation}
Where the true separation distance $D = \\sqrt{(Z_{hs} - Z_{gas})^2 + C_{lat}^2}$.

\\textbf{Parallel Orientation ($T_{para}$):}
\\begin{equation}
T_{para} = T_{surface} + \\frac{Q}{2\\pi k_{soil}} \\ln\\left(\\frac{d_{image}}{d_{source}}\\right)
\end{equation}
Where $d_{source} = \\sqrt{(Z_{hs} - Z_{gas})^2 + C^2}$ and $d_{image} = \\sqrt{(Z_{hs} + Z_{gas})^2 + C^2}$.

\\textbf{Blended Model for Angled Crossings ($0^{\\circ} < \\theta < 90^{\\circ}$):}
\\begin{equation}
T_{gas,homo} = w \\cdot T_{perp} + (1-w) \\cdot T_{para}, \\quad \\text{where } w = \\sin(\\theta)
\end{equation}

\\subsubsection{Gas Line Bedding Correction ($\\Delta T_{bed}$)}
A correction is applied if the gas line bedding material differs from the native soil.
\\begin{equation}
\\Delta T_{bed} = \\frac{Q}{2\\pi} \\ln\\left(\\frac{r_{b}}{r_{g}}\\right) \\left(\\frac{1}{k_{bed,gas}} - \\frac{1}{k_{soil}}\\right)
\end{equation}
The final temperature is then $T_{gas,layered} = T_{gas,homo} + \\Delta T_{bed}$.

\\subsubsection{Ground Surface Temperature ($T_{surf,hs}$)}
The temperature at y=1 inch below the surface, directly above the heat source.
\\begin{equation}
T_{surf,hs} = T_{surface} + \\frac{Q}{2\\pi k_{soil}} \\ln\\left(\\frac{y + Z_{hs}}{Z_{hs} - y}\\right)
\end{equation}

\\subsection{Variable Definitions}
\\begin{longtable}{l p{0.5\\textwidth} l}
\\toprule
\\textbf{Symbol} & \\textbf{Description} & \\textbf{Units} \\\\
\\midrule
\\endhead
\\bottomrule
\\endfoot
$Q$ & Heat loss per unit length & BTU/hr$\\cdot$ft \\\\
$T_{hs}$ & Temperature of heat source fluid & $^{\\circ}$F \\\\
$T_{surface}$ & Average ambient ground temperature & $^{\\circ}$F \\\\
$T_{gas}$ & Temperature at gas line & $^{\\circ}$F \\\\
$R$ & Thermal resistance per unit length & hr$\\cdot$ft$\\cdot$$^{\\circ}$F/BTU \\\\
$k$ & Thermal conductivity & BTU/hr$\\cdot$ft$\\cdot$$^{\\circ}$F \\\\
$D$ & Diameter & ft \\\\
$r$ & Radius & ft \\\\
$Z$ & Depth to centerline from ground surface & ft \\\\
$C$ & Centerline horizontal separation (parallel) & ft \\\\
$C_{lat}$ & Lateral offset at crossing point (perpendicular) & ft \\\\
$D$ & True centerline separation at crossing & ft \\\\
$d_{source}$ & Direct distance between pipes (parallel) & ft \\\\
$d_{image}$ & Distance from gas line to heat source's thermal image & ft \\\\
$\\theta$ & Crossing angle (90$^{\\circ}$ = perpendicular) & degrees \\\\
$y$ & Depth for surface temp calc (1 inch) & ft \\\\
\\bottomrule
\\caption{Subscripts: hs=heat source, gas=gas line, ins=insulation, bed=bedding, o=outer, i=inner, homo=homogeneous, perp=perpendicular, para=parallel} \\\\
\\end{longtable}
`;

    const createCalcWalkthrough = (results: any, title: string) => {
        if (!results) return '';
        const { R_pipe_wall, R_insulation, R_bedding_hs, R_soil_hs, R_total, Q, T_gas_line, T_gas_line_layered, T_ground_surface_above_hs, separation_distance, orientation_formula_used, inputs } = results;

        return `
\\subsubsection{${title}}
This section details the step-by-step calculation based on the input parameters for this scenario.

\\textbf{1. Thermal Resistances (hr$\\cdot$ft$\\cdot$$^{\\circ}$F/BTU)}
The total thermal resistance from the heat source fluid to the ambient soil is the sum of the resistances of each layer.

\\begin{itemize}[leftmargin=*]
    \\item $R_{pipe}$: Heat source pipe wall resistance.
    \\begin{align*}
        R_{pipe} &= \\frac{\\ln(${inputs.D_hs_outer.toFixed(3)} / ${inputs.D_hs_inner.toFixed(3)})}{2\\pi \\times ${inputs.k_pipe.toFixed(2)}} = \\mathbf{${R_pipe_wall.toFixed(4)}}
    \\end{align*}
    \\item $R_{ins}$: Insulation resistance.
    \\begin{align*}
        R_{ins} &= ${R_insulation > 0 ? `\\frac{\\ln(${inputs.D_ins_outer.toFixed(3)} / ${inputs.D_hs_outer.toFixed(3)})}{2\\pi \\times ${inputs.k_ins.toFixed(4)}} = \\mathbf{${R_insulation.toFixed(4)}}` : `\\mathbf{0.0000} \\text{ (No insulation in this scenario)}`}
    \\end{align*}
    \\item $R_{bed,hs}$: Heat source bedding resistance.
    \\begin{align*}
        R_{bed,hs} &= ${R_bedding_hs > 0 ? `\\frac{\\ln(${inputs.D_bed_outer.toFixed(3)} / ${inputs.D_ins_outer.toFixed(3)})}{2\\pi \\times ${inputs.k_bed_hs.toFixed(2)}} = \\mathbf{${R_bedding_hs.toFixed(4)}}` : `\\mathbf{0.0000} \\text{ (No bedding specified)}`}
    \\end{align*}
    \\item $R_{soil}$: Surrounding soil resistance, using the method of images.
    \\begin{align*}
        R_{soil} &= \\frac{\\ln((2 \\times ${inputs.Z_hs.toFixed(2)} - ${inputs.r_soil_i.toFixed(3)}) / ${inputs.r_soil_i.toFixed(3)})}{2\\pi \\times ${inputs.k_soil.toFixed(2)}} = \\mathbf{${R_soil_hs.toFixed(4)}}
    \\end{align*}
    \\item \\textbf{$R_{total}$} = $${R_pipe_wall.toFixed(4)} + ${R_insulation.toFixed(4)} + ${R_bedding_hs.toFixed(4)} + ${R_soil_hs.toFixed(4)} = \\mathbf{${R_total.toFixed(4)}}$
\\end{itemize}

\\textbf{2. Heat Loss (BTU/hr$\\cdot$ft)}
Using the total resistance, the heat loss per unit length of the source is calculated.
\\begin{align*}
    Q &= \\frac{T_{hs} - T_{surface}}{R_{total}} = \\frac{${inputs.T_hs.toFixed(1)} - ${inputs.T_surface.toFixed(1)}}{${R_total.toFixed(4)}} = \\mathbf{${Q.toFixed(2)}}
\\end{align*}

\\textbf{3. Gas Line Temperature - Homogeneous Soil ($T_{gas,homo}$)}
The temperature of the gas line is first calculated assuming it is surrounded by native soil.
\\begin{itemize}[leftmargin=*]
    \\item Geometry: Based on a \\textbf{${escapeLatex(orientation_formula_used)}} model, the effective separation distance is \\textbf{${separation_distance.toFixed(2)} ft}.
    \\item Temperature Rise:
    \\begin{align*}
      \\Delta T_{soil} &= T_{gas,homo} - T_{surface} = \\mathbf{${(T_gas_line - inputs.T_surface).toFixed(1)}}~^{\\circ}\\text{F}
    \\end{align*}
    \\item Homogeneous Temperature:
    \\begin{align*}
        T_{gas,homo} &= T_{surface} + \\Delta T_{soil} = ${inputs.T_surface.toFixed(1)} + ${(T_gas_line - inputs.T_surface).toFixed(1)} = \\mathbf{${T_gas_line.toFixed(1)}}~^{\\circ}\\text{F}
    \\end{align*}
\\end{itemize}

${!isNaN(T_gas_line_layered) ? `
\\textbf{4. Gas Line Bedding Correction ($\\Delta T_{bed}$)}
A correction is applied to account for the different thermal conductivity of the gas line bedding.
\\begin{itemize}[leftmargin=*]
    \\item Bedding Correction:
    \\begin{align*}
        \\Delta T_{bed} &= \\frac{${Q.toFixed(2)}}{2\\pi} \\ln\\left(\\frac{${inputs.r_b.toFixed(3)}}{${inputs.r_g.toFixed(3)}}\\right) \\left(\\frac{1}{${inputs.k_bed_gas.toFixed(2)}} - \\frac{1}{${inputs.k_soil.toFixed(2)}}\\right) = \\mathbf{${(T_gas_line_layered - T_gas_line).toFixed(1)}}~^{\\circ}\\text{F}
    \\end{align*}
\\end{itemize}
` : ''}

\\textbf{5. Final Temperatures ($^{\\circ}$F)}
\\begin{itemize}[leftmargin=*]
${!isNaN(T_gas_line_layered) ? `
    \\item \\textbf{Final Gas Line Temp (Layered Soil)}, $T_{gas,layered} = T_{gas,homo} + \\Delta T_{bed} = ${T_gas_line.toFixed(1)} + ${(T_gas_line_layered - T_gas_line).toFixed(1)} = \\mathbf{${T_gas_line_layered.toFixed(1)}}$
` : `
    \\item \\textbf{Final Gas Line Temp (Homogeneous Soil)}, $T_{gas,homo} = \\mathbf{${T_gas_line.toFixed(1)}}$
`}
    \\item Ground Surface Temp (1" deep), $T_{surf,hs} = \\mathbf{${T_ground_surface_above_hs.toFixed(1)}}$
\\end{itemize}
`;
    };

    const calcSection = `\\subsection{Calculation Walkthrough}
${createCalcWalkthrough(asIsResults, 'As-Is Scenario')}
${worstCaseResults ? `\\vspace{1cm} ${createCalcWalkthrough(worstCaseResults, 'Worst-Case Scenario (Insulation Failure)')}` : ''}
`;

    const gasLineTempRows = !isNaN(asIsResults.T_gas_line_layered)
        ? `Gas Line Temp (Homogeneous) ($^{\\circ}$F) & ${asIsResults.T_gas_line.toFixed(1)} & ${worstCaseResults ? worstCaseResults.T_gas_line.toFixed(1) : ''} \\\\
\\textbf{Gas Line Temp (with Bedding) ($^{\\circ}$F)} & \\textbf{${asIsResults.T_gas_line_layered.toFixed(1)}} & ${worstCaseResults && !isNaN(worstCaseResults.T_gas_line_layered) ? `\\textbf{${worstCaseResults.T_gas_line_layered.toFixed(1)}}` : ''}`
        : `\\textbf{Gas Line Temp ($^{\\circ}$F)} & \\textbf{${asIsResults.T_gas_line.toFixed(1)}} & ${worstCaseResults ? `\\textbf{${worstCaseResults.T_gas_line.toFixed(1)}}` : ''}`;

    const resultsSummary = `\\subsection{Results Summary}
\\begin{table}[h!]
\\centering
\\caption{Summary of Calculated Thermal Analysis Results}
\\begin{tabular}{lcc}
\\toprule
\\textbf{Parameter} & \\textbf{As-Is Scenario} & ${worstCaseResults ? '\\textbf{Worst-Case Scenario}' : ''} \\\\
\\midrule
Centerline Separation, D (ft) & ${asIsResults.separation_distance.toFixed(2)} & ${worstCaseResults ? worstCaseResults.separation_distance.toFixed(2) : ''} \\\\
${data.gasLineOrientation === 'Crossing / Perpendicular' ? `Crossing Angle (deg) & ${escapeLatex(data.crossingAngle)} & ${worstCaseResults ? escapeLatex(data.crossingAngle) : ''} \\\\` : ''}
Heat Loss, Q (BTU/hr$\\cdot$ft) & ${asIsResults.Q.toFixed(2)} & ${worstCaseResults ? worstCaseResults.Q.toFixed(2) : ''} \\\\
${gasLineTempRows} \\\\
Surface Temp, 1" deep ($^{\\circ}$F) & ${asIsResults.T_ground_surface_above_hs.toFixed(1)} & ${worstCaseResults ? worstCaseResults.T_ground_surface_above_hs.toFixed(1) : ''} \\\\
\\bottomrule
\\end{tabular}
\\end{table}
`;

    const recommendations = `\\section{Recommendations} \\label{sec:recommendations}
${(data.recommendations && data.recommendations.filter(r => r.trim()).length > 0) ? 
`\\begin{enumerate}[label=\\arabic*.]
${data.recommendations.filter(r => r.trim()).map(rec => `    \\item ${escapeLatex(rec)}`).join('\n')}
\\end{enumerate}` : 
'No specific recommendations were provided.'
}
`;

    const conclusion = `\\section{Conclusion}
This report provides a screening-level assessment based on analytical models. The results are subject to the assumptions and limitations outlined in the analysis, including the assumption of homogeneous soil and steady-state conditions. For critical applications, further analysis using numerical methods like Finite Element Analysis (FEA) may be warranted.
`;

    const endDoc = `\\end{document}`;
    
    return preamble + beginDoc + summary + inputData + analysis + calcSection + resultsSummary + recommendations + conclusion + endDoc;
};

document.addEventListener('DOMContentLoaded', () => {
  let isAdminAuthenticated = false;
  let targetTextareaIdForReword: string | null = null;
  let recommendationCounter = 0;

  const getPhotosData = () => {
    const photoItems = document.querySelectorAll('.photo-item');
    return Array.from(photoItems).map(item => {
        const img = item.querySelector('img');
        // FIX: The textarea for a photo description should have its own class to avoid conflicts.
        const textarea = item.querySelector<HTMLTextAreaElement>('.photo-description');
        return {
            src: img?.src || '',
            description: textarea?.value || ''
        };
    });
  };

  // --- Final Report Questionnaire State ---
  const reportQuestions: string[] = [
    "Are there any historical records of leaks, repairs, or integrity failures on either the heat source line or the gas line in the vicinity of this assessment?",
    "What is the expected remaining operational lifespan of the heat source system? Are there any scheduled integrity assessments, upgrades, or replacement plans?",
    "Considering the age and material of the gas line (e.g., plastic, coated/uncoated steel), what is its known susceptibility to temperature-related degradation, such as reduced tensile strength, embrittlement, or coating disbondment?",
    "Are there other underground utilities (e.g., water, sewer, electrical conduits) in close proximity that could be affected by the heat source or influence the heat transfer pathways to the gas line?",
    "What is the nature of the surface load in this area (e.g., roadway, landscaped area, sidewalk)? Could future excavation or changes in surface cover alter the current thermal conditions?",
    "Is the ground in this area prone to significant seasonal changes, such as deep frost penetration in winter or extreme drying in summer, that could alter soil thermal conductivity throughout the year?",
    "Are there any third-party construction projects planned nearby that could pose a risk of accidental damage to either the heat source or the gas line?",
    "What are the specific operational and safety consequences of the gas line's temperature exceeding its maximum allowable limit (e.g., pressure de-rating, immediate shutdown, public evacuation protocols)?",
    "Does the gas line operator have any specific guidelines, standards, or temperature thresholds for gas lines exposed to external heat sources that go beyond general regulatory requirements?",
    "Based on the preliminary findings, are there any immediately apparent, low-cost mitigation strategies that could be considered (e.g., installing monitoring equipment, improving surface drainage, adding warning markers)?",
    "In accordance with 49 CFR 192.317, which requires protecting gas pipelines from accidental damage, what specific physical barriers, markers, or shielding are recommended to prevent future excavation-related incidents in this area?",
    "Considering ASME B31.8's emphasis on material temperature limitations, if the calculated gas line temperature approaches the material's maximum allowable operating temperature (e.g., 140°F for PE), what specific de-rating of the MAOP (Maximum Allowable Operating Pressure) or other operational adjustments are recommended?",
    "Based on GTPC guidance for external heat sources, what long-term monitoring solutions (e.g., temperature sensors, periodic patrols with thermal cameras) are recommended to ensure the gas line's temperature remains within safe operational limits?",
    "Pursuant to 49 CFR 192.475, which covers the external corrosion control and the integrity of pipeline coatings, what recommendations are there for inspecting the gas line's coating in the heat-affected zone for potential disbondment or degradation?",
    "In alignment with the principles of pipeline integrity management (ASME B31.8S), what recommendations are there for updating the operator's integrity management plan (IMP) to include this specific heat source threat, including reassessment intervals and data integration?",
    "Other Recommendations: Please provide any other recommendations for mitigation, monitoring, or further analysis that have not been covered."
  ];
  let currentQuestionIndex = -1;
  let questionAnswers: { question: string; answer: string; originalQuestion: string }[] = [];
  let lastCalculationResults: any = null;

  // --- FIX: Add missing functions for Final Report tab ---
  const displayQuestion = () => {
    const questionTextEl = document.getElementById('question-text');
    const answerTextarea = document.getElementById('question-answer') as HTMLTextAreaElement;
    const progressText = document.getElementById('question-progress-text');
    const progressBar = document.getElementById('progress-bar');
    const nextBtn = document.getElementById('next-question-btn') as HTMLButtonElement;
    const prevBtn = document.getElementById('prev-question-btn') as HTMLButtonElement;
    
    if (!questionTextEl || !answerTextarea || !progressText || !progressBar || !nextBtn || !prevBtn) return;
    
    if (currentQuestionIndex >= 0 && currentQuestionIndex < reportQuestions.length) {
        const currentQ = reportQuestions[currentQuestionIndex];
        questionTextEl.textContent = currentQ;
        
        const savedAnswer = questionAnswers.find(qa => qa.originalQuestion === currentQ);
        answerTextarea.value = savedAnswer ? savedAnswer.answer : '';
        answerTextarea.dispatchEvent(new Event('input')); // Trigger auto-resize

        progressText.textContent = `Question ${currentQuestionIndex + 1} of ${reportQuestions.length}`;
        const progressPercentage = ((currentQuestionIndex + 1) / reportQuestions.length) * 100;
        progressBar.style.width = `${progressPercentage}%`;
        
        if (currentQuestionIndex === reportQuestions.length - 1) {
            nextBtn.textContent = 'Generate Report';
        } else {
            nextBtn.textContent = 'Next Question';
        }

        prevBtn.disabled = currentQuestionIndex === 0;
    }
  };

  const saveCurrentAnswer = () => {
    if (currentQuestionIndex < 0 || currentQuestionIndex >= reportQuestions.length) return;
    const currentQ = reportQuestions[currentQuestionIndex];
    const answer = (document.getElementById('question-answer') as HTMLTextAreaElement).value;
    
    const existingIndex = questionAnswers.findIndex(qa => qa.originalQuestion === currentQ);
    if (existingIndex > -1) {
      questionAnswers[existingIndex].answer = answer;
      questionAnswers[existingIndex].question = document.getElementById('question-text')?.textContent || currentQ;
    } else {
      questionAnswers.push({
        originalQuestion: currentQ,
        question: document.getElementById('question-text')?.textContent || currentQ,
        answer: answer
      });
    }
  };

  const generateReport = async () => {
    const container = document.getElementById('generated-report-container');
    const contentWrapper = document.getElementById('final-report-content');
    const loadingEl = document.getElementById('report-loading');

    if (!container || !contentWrapper || !loadingEl) return;
    
    if (!ai) {
        container.innerHTML = `<p class="error-message">AI client is not initialized. Please set a valid API key in the Admin tab to generate reports.</p>`;
        container.classList.remove('hidden');
        loadingEl.classList.add('hidden');
        contentWrapper.classList.remove('hidden'); // Show content wrapper again
        return;
    }
    
    contentWrapper.classList.add('hidden');
    loadingEl.classList.remove('hidden');

    try {
      const formData = getFormData();
      const photosData = getPhotosData();

      const photosForPrompt = photosData.map((p, index) => ({
        photoNumber: index + 1,
        description: p.description
      }));
      
      const prompt = `Generate a comprehensive engineering assessment report based on the following data.
The report should be well-structured with clear sections, headings, and bullet points, formatted as clean HTML.
It should identify potential risks based on the provided data and questionnaire answers, and provide actionable, specific recommendations for mitigation, monitoring, or further analysis.
The tone should be professional and technical, suitable for an engineering audience.
Start with a title and a summary section.

At the end of the report, create a section titled "Appendices", and within it, a subsection titled "Field Photos".
For each photo listed in the "Field Photos Data", create an entry in the report. The format for each photo entry should be:
<h3>Photo #[Photo Number]: [Photo Description]</h3>
[IMAGE_PLACEHOLDER_#[Photo Number]]

Project Data:
${JSON.stringify(formData, null, 2)}
      
Calculation Results (if available):
${JSON.stringify(lastCalculationResults, null, 2)}
      
Questionnaire Answers:
${JSON.stringify(questionAnswers, null, 2)}

Field Photos Data:
${photosData.length > 0 ? JSON.stringify(photosForPrompt, null, 2) : 'No photos were provided.'}
      `;

      const systemInstruction = `You are an expert engineering consultant specializing in pipeline integrity and thermal analysis for gas distribution systems. Your task is to generate a formal assessment report from the provided data.

When generating the final assessment report, you must adhere to the following strict operational limits and functional requirements:

**1. Pipeline Integrity Analysis Rules:**
*   **For All Plastic Pipe Types (e.g., HDPE, MDPE, Aldyl-A):**
    *   The maximum allowable operational temperature is **70°F**.
    *   **Reasoning:** Temperatures exceeding 70°F can result in a required de-rating of the pipe's Maximum Allowable Operating Pressure (MAOP) and can accelerate material degradation.
    *   **Required Action:** If the calculated gas line temperature meets or exceeds 70°F, you must state that this limit has been exceeded and recommend that a separate, detailed evaluation be conducted to assess the impact on MAOP and long-term pipe integrity.
*   **For All Steel Pipe Types (Coated, Unprotected Coated, and Bare):**
    *   The maximum allowable temperature is **150°F**.
    *   **Reasoning:** Temperatures greater than 150°F can damage the adhesive that bonds the protective coating to the steel pipe, potentially leading to coating disbondment and an increased risk of corrosion.
    *   **Required Action:** If the calculated gas line temperature meets or exceeds 150°F, you must state that this limit has been exceeded and recommend that a separate, detailed evaluation of the coating and pipeline integrity be conducted.

**2. Report Functionality Requirements:**
Editable Report Sections: When generating the final HTML report, ensure that all primary text-based sections (e.g., Executive Summary, Recommendations, Field Observations, Conclusion) are editable by the user directly in the browser. Achieve this by adding the contenteditable="true" attribute to the main container elements (e.g., <p>, <li>, <div>) for these sections.
PDF Export Capability: At the top of the generated report, include an action button labeled "Save as PDF". This button, when clicked, must execute the window.print() JavaScript function. Accompany the button with a brief instruction, such as: "To save as a PDF, click the button and select 'Save as PDF' from your browser's print destination options." You must also embed CSS within a <style> tag in the report's HTML to hide the button and this instruction during the printing process using an @media print query.

Important Constraint: Your sole function is to provide analysis and recommendations based on these rules. Do not suggest, perform, or otherwise indicate any changes to the assessment tool, its code, its user interface, or its underlying calculations. Your scope is strictly limited to the analytical task described.`;

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
        config: {
          systemInstruction,
        }
      });
      
      let reportHtml = response.text || '';

      // FIX: The model may wrap the HTML output in markdown fences (e.g., ```html...```).
      // This cleans the response to ensure it can be rendered as proper HTML.
      if (reportHtml.startsWith('```html')) {
        reportHtml = reportHtml.substring('```html'.length, reportHtml.length - 3).trim();
      } else if (reportHtml.startsWith('```')) {
        reportHtml = reportHtml.substring(3, reportHtml.length - 3).trim();
      }
      
      if (photosData.length > 0) {
        photosData.forEach((photo, index) => {
            if (photo.src) {
                const placeholder = `[IMAGE_PLACEHOLDER_#${index + 1}]`;
                // Sanitize description for alt text
                const sanitizedDescription = photo.description.replace(/"/g, '&quot;');
                const imgTag = `<img src="${photo.src}" alt="${sanitizedDescription}" style="max-width: 100%; height: auto; border: 1px solid #ccc; margin-top: 0.5rem;">`;
                const regex = new RegExp(placeholder.replace(/\[/g, "\\[").replace(/\]/g, "\\]").replace(/#/g, "\\#"), "g");
                reportHtml = reportHtml.replace(regex, imgTag);
            }
        });
      }

      container.innerHTML = reportHtml;

      // --- FIX: Override the default print functionality with a robust PDF generator ---
      const pdfButton = Array.from(container.querySelectorAll('button')).find(
        btn => btn.textContent?.trim() === 'Save as PDF'
      );

      if (pdfButton) {
        // Hijack the button's click event for a better PDF export.
        // This replaces the model's default onclick="window.print()" which can be unreliable.
        pdfButton.removeAttribute('onclick');
        pdfButton.addEventListener('click', () => {
          const originalButtonText = pdfButton.textContent;
          pdfButton.textContent = 'Generating PDF...';
          pdfButton.disabled = true;
          
          // Find the button's container to hide it and any accompanying instructions.
          const buttonContainer = pdfButton.parentElement;
          if (buttonContainer) {
            buttonContainer.style.display = 'none';
          }

          const opt = {
            margin:       [0.5, 0.5, 0.5, 0.5],
            filename:     `final-report-${new Date().toISOString().split('T')[0]}.pdf`,
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2, useCORS: true, letterRendering: true, scrollY: 0 },
            jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' },
            pagebreak:    { mode: ['avoid-all', 'css', 'legacy'] }
          };

          (window as any).html2pdf().from(container).set(opt).save().then(() => {
            pdfButton.textContent = originalButtonText;
            pdfButton.disabled = false;
            if (buttonContainer) {
              buttonContainer.style.display = '';
            }
          }).catch((err: Error) => {
              console.error("Failed to export Final Report PDF:", err);
              alert("An error occurred during PDF export. Please check the console for details.");
              pdfButton.textContent = originalButtonText;
              pdfButton.disabled = false;
              if (buttonContainer) {
                buttonContainer.style.display = '';
              }
          });
        });
      }

      container.classList.remove('hidden');
      // Re-run the tab handler to add the "Generate New" button to the new report
      handleFinalReportTabClick();

    } catch(error) {
      console.error("Report Generation Error:", error);
      container.innerHTML = `<p class="error-message">Failed to generate report. Please check your API key and the browser console for details. Error: ${error instanceof Error ? error.message : 'Unknown error'}</p>`;
      container.classList.remove('hidden');
    } finally {
      loadingEl.classList.add('hidden');
    }
  };

  const handleNextQuestion = async () => {
    saveCurrentAnswer();
    
    if (currentQuestionIndex < reportQuestions.length - 1) {
      currentQuestionIndex++;
      displayQuestion();
    } else {
      await generateReport();
    }
  };

  const handlePrevQuestion = () => {
    saveCurrentAnswer();

    if (currentQuestionIndex > 0) {
      currentQuestionIndex--;
      displayQuestion();
    }
  };

  const startQuestionnaire = () => {
    const startWrapper = document.getElementById('final-report-start-wrapper');
    const questionnaireWrapper = document.getElementById('final-report-questionnaire-wrapper');
    
    if (startWrapper && questionnaireWrapper) {
      startWrapper.classList.add('hidden');
      questionnaireWrapper.classList.remove('hidden');
      currentQuestionIndex = 0;
      displayQuestion();
    }
  };

  const handleFinalReportTabClick = () => {
    const adminMessage = document.getElementById('report-admin-only-message');
    const contentWrapper = document.getElementById('final-report-content');
    const startWrapper = document.getElementById('final-report-start-wrapper');
    const questionnaireWrapper = document.getElementById('final-report-questionnaire-wrapper');
    const generatedReportContainer = document.getElementById('generated-report-container');
    const loadingEl = document.getElementById('report-loading');

    if (!adminMessage || !contentWrapper || !startWrapper || !questionnaireWrapper || !generatedReportContainer || !loadingEl) return;

    adminMessage.classList.toggle('hidden', isAdminAuthenticated);
    contentWrapper.classList.toggle('hidden', !isAdminAuthenticated);
    
    if (!isAdminAuthenticated) {
        return;
    }

    // If a report is already displayed, add a "generate new" button to it.
    if (!generatedReportContainer.classList.contains('hidden')) {
        if (!document.getElementById('report-actions-wrapper')) {
            const actionsWrapper = document.createElement('div');
            actionsWrapper.id = 'report-actions-wrapper';
            const regenerateButton = document.createElement('button');
            regenerateButton.textContent = 'Generate New Report';
            regenerateButton.className = 'action-button';
            regenerateButton.onclick = () => {
                // Reset state
                generatedReportContainer.classList.add('hidden');
                generatedReportContainer.innerHTML = '';
                currentQuestionIndex = -1;
                questionAnswers = [];
                // Re-run this handler to show the start options
                handleFinalReportTabClick();
            };
            actionsWrapper.appendChild(regenerateButton);
            generatedReportContainer.prepend(actionsWrapper);
        }
        return; // Don't proceed to show other things
    }
    
    // If loading, do nothing
    if (!loadingEl.classList.contains('hidden')) {
        return;
    }

    if (!lastCalculationResults) {
        startWrapper.innerHTML = `<p>Please run a calculation on the "Calculation" tab before generating a report.</p>`;
        startWrapper.classList.remove('hidden');
        questionnaireWrapper.classList.add('hidden');
    } else {
        if (currentQuestionIndex < 0) {
            // Questionnaire has not been started yet, show the generation options.
            startWrapper.innerHTML = `
                <p>The calculation has been completed. You can now generate the final assessment report.</p>
                <div class="form-group-info" style="margin-top: 1rem;">
                    You can generate a report directly, or answer guided questions to add more detail for a more thorough analysis.
                </div>
                <div class="report-generation-options">
                    <button id="generateQuickReportButton" class="action-button">Generate Report Now</button>
                    <button id="startQuestionnaireButton" class="action-button secondary">Start Guided Report</button>
                </div>
            `;
            document.getElementById('startQuestionnaireButton')?.addEventListener('click', startQuestionnaire);
            document.getElementById('generateQuickReportButton')?.addEventListener('click', () => {
                questionAnswers = []; // Clear answers for a non-guided report
                generateReport();
            });

            startWrapper.classList.remove('hidden');
            questionnaireWrapper.classList.add('hidden');
        } else {
            // Questionnaire is already in progress, just show it.
            startWrapper.classList.add('hidden');
            questionnaireWrapper.classList.remove('hidden');
        }
    }
  };
  // --- END OF FIX ---


  // --- Tab Navigation ---
  const tabButtons = document.querySelectorAll('.tab-button');
  const tabPanels = document.querySelectorAll('.tab-panel');

  tabButtons.forEach(button => {
    button.addEventListener('click', () => {
      // Deactivate all tabs and panels
      tabButtons.forEach(btn => btn.classList.remove('active'));
      tabPanels.forEach(panel => panel.classList.remove('active'));

      // Activate the clicked tab and corresponding panel
      const tabId = button.getAttribute('data-tab');
      if (tabId === 'calculation') {
        updateCalculationSummary();
      }
      
      if (tabId === 'assessment-form') {
        updateAssessmentFormView();
      }

      if (tabId === 'admin') {
        if (isAdminAuthenticated) {
            document.getElementById('admin-login-wrapper')?.classList.add('hidden');
            document.getElementById('admin-content-wrapper')?.classList.remove('hidden');
        } else {
            document.getElementById('admin-login-wrapper')?.classList.remove('hidden');
            document.getElementById('admin-content-wrapper')?.classList.add('hidden');
        }
      }

      if (tabId === 'final-report') {
        handleFinalReportTabClick();
      }

      const targetPanel = document.getElementById(tabId!);
      button.classList.add('active');
      if (targetPanel) {
        targetPanel.classList.add('active');
      }
    });
  });

  // --- Dynamic Evaluator Inputs ---
  const numEvaluatorsSelect = document.getElementById('numEvaluators') as HTMLSelectElement;
  const evaluatorNamesContainer = document.getElementById('evaluatorNamesContainer');

  const updateEvaluatorInputs = () => {
      if (!numEvaluatorsSelect || !evaluatorNamesContainer) return;
      const count = parseInt(numEvaluatorsSelect.value, 10);
      evaluatorNamesContainer.innerHTML = ''; // Clear existing inputs
      for (let i = 1; i <= count; i++) {
          const input = document.createElement('input');
          input.type = 'text';
          input.className = 'evaluator-name-input';
          input.name = `evaluatorName${i}`;
          input.placeholder = `Enter name of evaluator ${i}`;
          evaluatorNamesContainer.appendChild(input);
      }
  };

  numEvaluatorsSelect?.addEventListener('change', updateEvaluatorInputs);


  // --- Helper for Textarea Auto-Resize ---
  const enableTextareaAutoResize = (elementId: string) => {
    const textarea = document.getElementById(elementId) as HTMLTextAreaElement;
    if (textarea) {
      const adjustTextareaHeight = () => {
        textarea.style.height = 'auto';
        textarea.style.height = `${textarea.scrollHeight}px`;
      };
      textarea.addEventListener('input', adjustTextareaHeight);
      adjustTextareaHeight(); // Initial call
    }
  };

  // Enable auto-resize for textareas
  enableTextareaAutoResize('projectDescription');
  enableTextareaAutoResize('additionalInfo');
  enableTextareaAutoResize('customConnectionTypes');
  enableTextareaAutoResize('systemDutyCycle');
  enableTextareaAutoResize('pipeCasingInfo');
  enableTextareaAutoResize('heatLossEvidence');
  enableTextareaAutoResize('insulationCondition');
  enableTextareaAutoResize('waterInfiltrationComments');
  enableTextareaAutoResize('question-answer');
  enableTextareaAutoResize('sitePersonnel');
  enableTextareaAutoResize('siteConditions');
  enableTextareaAutoResize('fieldObservations');


  // --- Conditional display for the Heat Source Data tab ---
  const heatSourceTypeRadios = document.querySelectorAll<HTMLInputElement>('input[name="heatSourceType"]');
  const conditionalFields = document.getElementById('conditionalFields');
  const heatSourceHeader = document.getElementById('heatSourceHeader') as HTMLHeadingElement;
  const operatorNameLabel = document.getElementById('operatorNameLabel') as HTMLLabelElement;
  const operatorCompanyNameLabel = document.getElementById('operatorCompanyNameLabel') as HTMLLabelElement;
  const operatorCompanyAddressLabel = document.getElementById('operatorCompanyAddressLabel') as HTMLLabelElement;
  const operatorContactInfoLabel = document.getElementById('operatorContactInfoLabel') as HTMLLabelElement;
  const confirmationDateLabel = document.getElementById('confirmationDateLabel') as HTMLLabelElement;
  const maxTempLabel = document.getElementById('maxTempLabel') as HTMLLabelElement;
  const maxPressureLabel = document.getElementById('maxPressureLabel') as HTMLLabelElement;
  const heatSourceAgeLabel = document.getElementById('heatSourceAgeLabel') as HTMLLabelElement;
  const systemDutyCycleLabel = document.getElementById('systemDutyCycleLabel') as HTMLLabelElement;
  const casingInfoLabel = document.getElementById('casingInfoLabel') as HTMLLabelElement;
  const pipelineDiameterLabel = document.getElementById('pipelineDiameterLabel') as HTMLLabelElement;
  const pipeMaterialLabel = document.getElementById('pipeMaterialLabel') as HTMLLabelElement;
  const wallThicknessLabel = document.getElementById('wallThicknessLabel') as HTMLLabelElement;
  const connectionTypesLabel = document.getElementById('connectionTypesLabel') as HTMLLabelElement;
  const customPipeMaterialLabel = document.getElementById('customPipeMaterialLabel') as HTMLLabelElement;
  const customThermalConductivityLabel = document.getElementById('customThermalConductivityLabel') as HTMLLabelElement;
  const pipeInsulationTypeLabel = document.getElementById('pipeInsulationTypeLabel') as HTMLLabelElement;
  const customPipeInsulationLabel = document.getElementById('customPipeInsulationLabel') as HTMLLabelElement;
  const customInsulationThermalConductivityLabel = document.getElementById('customInsulationThermalConductivityLabel') as HTMLLabelElement;
  const insulationThicknessLabel = document.getElementById('insulationThicknessLabel') as HTMLLabelElement;
  const additionalInfoLabel = document.getElementById('additionalInfoLabel') as HTMLLabelElement;
  const heatSourceDepthLabel = document.getElementById('heatSourceDepthLabel') as HTMLLabelElement;

  const toggleConditionalFields = () => {
    const selectedRadio = document.querySelector<HTMLInputElement>('input[name="heatSourceType"]:checked');
    const selectedValue = selectedRadio?.value;
    const customHeatSourceWrapper = document.getElementById('customHeatSourceWrapper');

    if (customHeatSourceWrapper) {
      customHeatSourceWrapper.classList.toggle('hidden', selectedValue !== 'other');
    }

    let sourceName = 'Heat Source';
    if (selectedRadio) {
      const label = document.querySelector(`label[for="${selectedRadio.id}"]`);
      if (label && label.textContent) {
        if (selectedValue === 'other') {
          const customInput = document.getElementById('customHeatSourceType') as HTMLInputElement;
          sourceName = customInput?.value.trim() || 'Other Heat Source';
        } else {
          sourceName = label.textContent.trim();
        }
      }
    }
    
    // Dynamically update labels based on selected heat source
    if (heatSourceHeader) heatSourceHeader.textContent = `${sourceName} Data`;
    if (operatorNameLabel) operatorNameLabel.textContent = `${sourceName} - Name of Individual(s) Providing Data:`;
    if (operatorCompanyNameLabel) operatorCompanyNameLabel.textContent = `${sourceName} - Operator Company Name:`;
    if (operatorCompanyAddressLabel) operatorCompanyAddressLabel.textContent = `${sourceName} - Operating Company Address:`;
    if (operatorContactInfoLabel) operatorContactInfoLabel.textContent = `${sourceName} - Contact Info (Phone or Email):`;
    if (confirmationDateLabel) confirmationDateLabel.textContent = `${sourceName} Data Confirmation Date:`;
    if (maxTempLabel) maxTempLabel.textContent = `${sourceName} Max Operating Temp (°F) confirmed by the operator:`;
    if (maxPressureLabel) maxPressureLabel.textContent = `${sourceName} Max Operating Pressure in psig:`;
    if (heatSourceAgeLabel) heatSourceAgeLabel.textContent = `${sourceName} Line Age (years):`;
    if (systemDutyCycleLabel) systemDutyCycleLabel.textContent = `${sourceName} System Duty Cycle / Uptime:`;
    if (casingInfoLabel) casingInfoLabel.textContent = `${sourceName} Pipe Casing / Conduit Information:`;
    if (pipelineDiameterLabel) pipelineDiameterLabel.textContent = `${sourceName} Pipeline Nominal Diameter (inches):`;
    if (pipeMaterialLabel) pipeMaterialLabel.textContent = `${sourceName} Pipe Material:`;
    if (wallThicknessLabel) wallThicknessLabel.textContent = `${sourceName} Pipe Wall Thickness (inches):`;
    if (connectionTypesLabel) {
      let connectionLabelSourceName = sourceName.replace(/ Line$/, ''); // "Steam Line" -> "Steam"
      connectionTypesLabel.textContent = `${connectionLabelSourceName} connection types:`;
    }
    if (customPipeMaterialLabel) customPipeMaterialLabel.textContent = `${sourceName} Custom Pipe Material Name:`;
    if (customThermalConductivityLabel) customThermalConductivityLabel.textContent = `${sourceName} Thermal Conductivity (BTU/hr·ft·°F):`;
    if (pipeInsulationTypeLabel) pipeInsulationTypeLabel.textContent = `${sourceName} Pipe Insulation Type:`;
    if (customPipeInsulationLabel) customPipeInsulationLabel.textContent = `${sourceName} Custom Pipe Insulation Name:`;
    if (customInsulationThermalConductivityLabel) customInsulationThermalConductivityLabel.textContent = `${sourceName} Thermal Conductivity (BTU/hr·ft·°F):`;
    if (insulationThicknessLabel) insulationThicknessLabel.textContent = `${sourceName} Insulation Thickness (inches):`;
    if (heatSourceDepthLabel) heatSourceDepthLabel.textContent = `${sourceName} Depth from Ground Surface to Pipe Centerline (feet):`;
    if (additionalInfoLabel) additionalInfoLabel.textContent = `${sourceName} - Additional Information Provided by the Operator:`;
    
    // Show/hide the main form section
    if (conditionalFields) {
      if (selectedValue === 'steam' || selectedValue === 'hotWater' || selectedValue === 'superHeatedHotWater' || selectedValue === 'other') {
        conditionalFields.classList.remove('hidden');
      } else {
        conditionalFields.classList.add('hidden');
      }
    }

    updateInsulationOptions(selectedValue);
  };

  heatSourceTypeRadios.forEach(radio => {
    radio.addEventListener('change', toggleConditionalFields);
  });
  document.getElementById('customHeatSourceType')?.addEventListener('input', toggleConditionalFields);

  // --- Data for future calculations ---
  const npsToOdMapping: { [key: string]: number } = {
    '1': 1.315,
    '1.25': 1.660,
    '1.5': 1.900,
    '2': 2.375,
    '3': 3.500,
    '4': 4.500,
    '5': 5.563,
    '6': 6.625,
    '8': 8.625,
    '10': 10.750,
    '12': 12.750,
    '14': 14.000,
    '16': 16.000,
    '18': 18.000,
    '20': 20.000,
    '22': 22.000,
    '24': 24.000,
    '26': 26.000,
    '30': 30.000,
    '36': 36.000
  };

  const getGasPipeOd = (nominalStr: string, standard: string): number => {
    if (!nominalStr || !standard || nominalStr === 'N/A') return NaN;

    const nominalNum = parseFloat(nominalStr);
    if (isNaN(nominalNum)) return NaN;

    if (standard === 'ips') {
      return npsToOdMapping[nominalStr] || nominalNum;
    } else if (standard === 'cts') {
      return nominalNum + 0.125;
    }
    return NaN;
  };

  const pipeMaterialData = {
    'carbon-steel': { name: 'Carbon Steel', thermalConductivity: 26 },
    'stainless-steel': { name: 'Stainless Steel', thermalConductivity: 9 },
    'ductile-iron': { name: 'Ductile Iron', thermalConductivity: 32 },
    'cast-iron': { name: 'Cast Iron (Legacy)', thermalConductivity: 27 },
    'copper': { name: 'Copper', thermalConductivity: 223 },
    'pex': { name: 'PEX (Hot Water Only)', thermalConductivity: 0.26 },
    'pp-r': { name: 'PP-R / PP-RCT (Hot Water Only)', thermalConductivity: 0.22 },
    'frp': { name: 'Fiberglass Reinforced Pipe (FRP)', thermalConductivity: 0.3 },
  };

  const insulationData: { [key: string]: { [key: string]: { name: string; thermalConductivity: number } } } = {
    steam: {
      'calcium-silicate': { name: 'Calcium Silicate', thermalConductivity: 0.034 },
      'mineral-wool': { name: 'Mineral Wool (Rock/Slag)', thermalConductivity: 0.023 },
      'fiberglass': { name: 'Fiberglass (Pipe Sections)', thermalConductivity: 0.021 },
      'cellular-glass': { name: 'Cellular Glass', thermalConductivity: 0.032 },
    },
    hotWater: {
      'polyurethane-foam': { name: 'Polyurethane Foam (PUR)', thermalConductivity: 0.015 },
      'polyisocyanurate-foam': { name: 'Polyisocyanurate Foam (PIR)', thermalConductivity: 0.016 },
      'phenolic-foam': { name: 'Phenolic Foam', thermalConductivity: 0.012 },
      'cellular-glass': { name: 'Cellular Glass', thermalConductivity: 0.032 },
      'fiberglass': { name: 'Fiberglass (Pipe Sections)', thermalConductivity: 0.021 },
    }
  };

  const pipeWallThicknessData: { [key: string]: number[] } = {
    '1': [0.109, 0.133, 0.179], '1.25': [0.109, 0.140, 0.191], '1.5': [0.109, 0.145, 0.200], '2': [0.109, 0.154, 0.218], '3': [0.120, 0.216, 0.300], '4': [0.120, 0.237, 0.337], '5': [0.258, 0.375], '6': [0.280, 0.432], '8': [0.250, 0.277, 0.322, 0.406, 0.500], '10': [0.250, 0.307, 0.365, 0.500, 0.594], '12': [0.250, 0.330, 0.375, 0.406, 0.500, 0.688], '14': [0.250, 0.312, 0.375, 0.438, 0.500], '16': [0.250, 0.312, 0.375, 0.500, 0.656], '18': [0.250, 0.312, 0.375, 0.500, 0.750], '20': [0.250, 0.375, 0.500, 0.594], '22': [0.250, 0.375, 0.500], '24': [0.250, 0.375, 0.500, 0.688], '26': [0.312, 0.500], '30': [0.312, 0.375, 0.500], '36': [0.312, 0.375, 0.500],
  };

  // --- Generic Helper for Diameter -> Wall Thickness Logic ---
  const createWallThicknessHandler = (diameterSelectId: string, thicknessSelectId: string, customWrapperId: string, customInputId: string) => {
    const diameterSelect = document.getElementById(diameterSelectId) as HTMLSelectElement;
    const thicknessSelect = document.getElementById(thicknessSelectId) as HTMLSelectElement;
    const customWrapper = document.getElementById(customWrapperId);
    const customInput = document.getElementById(customInputId) as HTMLInputElement;

    const updateOptions = () => {
      const selectedDiameter = diameterSelect.value;
      const thicknesses = pipeWallThicknessData[selectedDiameter] || [];
      const currentVal = thicknessSelect.value; // Save current value before clearing
      thicknessSelect.innerHTML = '';
      if (customWrapper) customWrapper.classList.add('hidden');
      
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = thicknesses.length > 0 ? 'Select thickness...' : 'No standard thicknesses';
      placeholder.disabled = true;
      placeholder.selected = true;
      thicknessSelect.appendChild(placeholder);
      
      let currentValExists = false;
      thicknesses.forEach(thickness => {
        const option = document.createElement('option');
        option.value = String(thickness);
        option.textContent = `${thickness}"`;
        if(String(thickness) === currentVal) {
          option.selected = true;
          currentValExists = true;
        }
        thicknessSelect.appendChild(option);
      });

      const customOption = document.createElement('option');
      customOption.value = 'custom';
      customOption.textContent = 'Other (specify)...';
      if (currentVal && !currentValExists && currentVal !== 'custom') {
          // If a custom value was loaded, keep custom selected and set the input
          customOption.selected = true;
          if (customInput) customInput.value = currentVal;
          if (customWrapper) customWrapper.classList.remove('hidden');
      } else if (currentVal === 'custom') {
          customOption.selected = true;
      }

      thicknessSelect.appendChild(customOption);
      thicknessSelect.disabled = false;
    };

    const handleThicknessChange = () => {
      if (thicknessSelect.value === 'custom' && customWrapper) {
        customWrapper.classList.remove('hidden');
      } else if (customWrapper) {
        customWrapper.classList.add('hidden');
      }
    };
    diameterSelect.addEventListener('change', updateOptions);
    thicknessSelect.addEventListener('change', handleThicknessChange);
  };
  
  createWallThicknessHandler('pipelineDiameter', 'wallThickness', 'customWallThicknessWrapper', 'customWallThickness');
  createWallThicknessHandler('gasPipelineDiameter', 'gasWallThickness', 'customGasWallThicknessWrapper', 'customGasWallThickness');


  // --- Generic Helper for Material -> Custom Material & Thermal Conductivity ---
  const createPipeMaterialHandler = (materialSelectId: string, customWrapperId: string, conductivityDisplayId: string) => {
    const materialSelect = document.getElementById(materialSelectId) as HTMLSelectElement;
    const customWrapper = document.getElementById(customWrapperId);
    const conductivityDisplay = document.getElementById(conductivityDisplayId);

    const handleChange = () => {
        if (!materialSelect) return;
        const selectedValue = materialSelect.value;
        const isOtherSelected = selectedValue === 'other';

        // Show/hide the custom input fields wrapper based on selection
        if (customWrapper) {
            customWrapper.classList.toggle('hidden', !isOtherSelected);
        }

        // Show/hide the pre-calculated thermal conductivity display
        if (conductivityDisplay) {
            if (isOtherSelected) {
                conductivityDisplay.classList.add('hidden');
            } else {
                const material = pipeMaterialData[selectedValue as keyof typeof pipeMaterialData];
                if (material) {
                    conductivityDisplay.textContent = `Thermal Conductivity: ${material.thermalConductivity} BTU/hr·ft·°F`;
                    conductivityDisplay.classList.remove('hidden');
                } else {
                    conductivityDisplay.classList.add('hidden');
                }
            }
        }
    };

    if (materialSelect) {
        materialSelect.addEventListener('change', handleChange);
        handleChange(); // Set initial state
    }
  };

  createPipeMaterialHandler('pipeMaterial', 'customPipeMaterialWrapper', 'thermalConductivityDisplay');
  createPipeMaterialHandler('gasPipeMaterial', 'customGasPipeMaterialWrapper', 'gasThermalConductivityDisplay');

  // --- Gas Line Operator Name Logic ---
  const gasOperatorNameSelect = document.getElementById('gasOperatorName') as HTMLSelectElement;
  const customGasOperatorNameWrapper = document.getElementById('customGasOperatorNameWrapper');

  const handleGasOperatorNameChange = () => {
    if (gasOperatorNameSelect.value === 'other') {
      customGasOperatorNameWrapper?.classList.remove('hidden');
    } else {
      customGasOperatorNameWrapper?.classList.add('hidden');
    }
  };
  gasOperatorNameSelect.addEventListener('change', handleGasOperatorNameChange);

  // --- Gas Line Material, Coating, Wall Thickness, and SDR Logic ---
  const gasPipeMaterialSelect = document.getElementById('gasPipeMaterial') as HTMLSelectElement;
  const gasCoatingTypeWrapper = document.getElementById('gasCoatingTypeWrapper');
  const gasCoatingTypeSelect = document.getElementById('gasCoatingType') as HTMLSelectElement;
  const customGasCoatingTypeWrapper = document.getElementById('customGasCoatingTypeWrapper');
  const customGasCoatingTypeInput = document.getElementById('customGasCoatingType') as HTMLInputElement;
  const gasWallThicknessGroup = document.getElementById('gasWallThicknessGroup');
  const gasSdrGroup = document.getElementById('gasSdrGroup');
  const gasSdrSelect = document.getElementById('gasSdr') as HTMLSelectElement;
  const customGasSdrWrapper = document.getElementById('customGasSdrWrapper');
  
  const handleGasCoatingTypeChange = () => {
    if (gasCoatingTypeSelect?.value === 'other') {
      customGasCoatingTypeWrapper?.classList.remove('hidden');
    } else {
      customGasCoatingTypeWrapper?.classList.add('hidden');
      if (customGasCoatingTypeInput) customGasCoatingTypeInput.value = '';
    }
  };
  
  const handleGasPipeMaterialChange = () => {
    const selectedValue = gasPipeMaterialSelect.value;
    const showWallThickness = ['coated-steel-protected', 'coated-steel-unprotected', 'bare-steel', 'pvc'].includes(selectedValue);
    const isPoly = ['hdpe', 'mdpe', 'aldyl'].includes(selectedValue);

    // Show/hide Wall Thickness for Steel or PVC
    gasWallThicknessGroup?.classList.toggle('hidden', !showWallThickness);

    // Show/hide SDR for Polyethylene
    gasSdrGroup?.classList.toggle('hidden', !isPoly);

    // Show/hide Coating Type for Coated Steel
    const showCoating = selectedValue === 'coated-steel-protected' || selectedValue === 'coated-steel-unprotected';
    gasCoatingTypeWrapper?.classList.toggle('hidden', !showCoating);
    if (!showCoating && gasCoatingTypeSelect) {
      gasCoatingTypeSelect.value = '';
      // Trigger the dependent dropdown's handler to ensure its state is correct
      handleGasCoatingTypeChange();
    }
  };
  
  const handleGasSdrChange = () => {
    if (customGasSdrWrapper) {
      customGasSdrWrapper.classList.toggle('hidden', gasSdrSelect.value !== 'other');
    }
  };

  if(gasPipeMaterialSelect) gasPipeMaterialSelect.addEventListener('change', handleGasPipeMaterialChange);
  if(gasCoatingTypeSelect) gasCoatingTypeSelect.addEventListener('change', handleGasCoatingTypeChange);
  if(gasSdrSelect) gasSdrSelect.addEventListener('change', handleGasSdrChange);


  // --- Gas Line Orientation Logic ---
  const gasLineOrientationRadios = document.querySelectorAll<HTMLInputElement>('input[name="gasLineOrientation"]');
  const parallelDistanceWrapper = document.getElementById('parallelDistanceWrapper');
  const perpendicularCoordinatesWrapper = document.getElementById('perpendicularCoordinatesWrapper');
  
  const handleGasLineOrientationChange = () => {
      const selectedRadio = document.querySelector<HTMLInputElement>('input[name="gasLineOrientation"]:checked');
      if (parallelDistanceWrapper) {
        parallelDistanceWrapper.classList.toggle('hidden', selectedRadio?.value !== 'parallel');
      }
      if (perpendicularCoordinatesWrapper) {
        perpendicularCoordinatesWrapper.classList.toggle('hidden', selectedRadio?.value !== 'perpendicular');
      }
  };
  
  gasLineOrientationRadios.forEach(radio => {
      radio.addEventListener('change', handleGasLineOrientationChange);
  });

  const updateParallelCoordinateInputs = () => {
    const lengthInput = document.getElementById('parallelLength') as HTMLInputElement;
    const container = document.getElementById('parallelCoordinatesContainer');
    const wrapper = document.getElementById('parallelCoordinatesWrapper');
    if (!lengthInput || !container || !wrapper) return;

    const length = parseFloat(lengthInput.value);
    wrapper.innerHTML = '';

    if (isNaN(length) || length <= 0) {
        container.classList.add('hidden');
        return;
    }

    container.classList.remove('hidden');
    
    wrapper.innerHTML = `
        <div class="parallel-coordinate-grid">
            <div class="grid-header">Point</div>
            <div class="grid-header">Latitude</div>
            <div class="grid-header">Longitude</div>
        </div>
    `;

    const createRow = (labelText: string, index: number) => {
        const grid = document.createElement('div');
        grid.className = 'parallel-coordinate-grid';

        const label = document.createElement('label');
        label.textContent = labelText;
        label.htmlFor = `lat-parallel-${index}`;

        const latInput = document.createElement('input');
        latInput.type = 'number';
        latInput.id = `lat-parallel-${index}`;
        latInput.placeholder = 'e.g., 42.3601';
        latInput.step = 'any';
        latInput.dataset.pointLabel = labelText;

        const lngInput = document.createElement('input');
        lngInput.type = 'number';
        lngInput.id = `lng-parallel-${index}`;
        lngInput.placeholder = 'e.g., -71.0589';
        lngInput.step = 'any';

        grid.appendChild(label);
        grid.appendChild(latInput);
        grid.appendChild(lngInput);
        
        return grid;
    };

    // Start point
    wrapper.appendChild(createRow('Start', 0));

    // Intermediate points
    let pointIndex = 1;
    for (let dist = 10; dist < length; dist += 10) {
        wrapper.appendChild(createRow(`${dist} ft`, pointIndex));
        pointIndex++;
    }
    
    // End point - ensure it's always added if length > 0
    wrapper.appendChild(createRow('End', pointIndex));
  };
  
  const parallelLengthInput = document.getElementById('parallelLength') as HTMLInputElement;
  parallelLengthInput?.addEventListener('input', updateParallelCoordinateInputs);

  // --- Heat Source Insulation Logic ---
  const pipeInsulationTypeSelect = document.getElementById('pipeInsulationType') as HTMLSelectElement;
  const insulationThicknessWrapper = document.getElementById('insulationThicknessWrapper');
  const customPipeInsulationWrapper = document.getElementById('customPipeInsulationWrapper');
  const insulationThermalConductivityDisplay = document.getElementById('insulationThermalConductivityDisplay');

  const updateInsulationOptions = (heatSourceType: string | undefined) => {
      // FIX: Preserve the current selection when options are repopulated to prevent it from being cleared by event triggers.
      const currentValue = pipeInsulationTypeSelect.value;
  
      let options: { [key: string]: { name: string; thermalConductivity: number } } = {};
      if (heatSourceType === 'steam' || heatSourceType === 'superHeatedHotWater') {
          options = insulationData.steam;
      } else if (heatSourceType === 'hotWater') {
          options = insulationData.hotWater;
      } else if (heatSourceType === 'other') {
          options = { ...insulationData.steam, ...insulationData.hotWater };
      }

      let bestCaseKey: string | null = null;
      let worstCaseKey: string | null = null;
      
      const optionKeys = Object.keys(options);
      if (optionKeys.length > 1) {
          let minConductivity = Infinity;
          let maxConductivity = -Infinity;

          for (const key in options) {
              const conductivity = options[key].thermalConductivity;
              if (conductivity < minConductivity) {
                  minConductivity = conductivity;
                  bestCaseKey = key;
              }
              if (conductivity > maxConductivity) {
                  maxConductivity = conductivity;
                  worstCaseKey = key;
              }
          }
          
          if (minConductivity === maxConductivity) {
              bestCaseKey = null;
              worstCaseKey = null;
          }
      }
      
      pipeInsulationTypeSelect.innerHTML = '';
      const placeholder = document.createElement('option');
      placeholder.value = '';
      placeholder.textContent = Object.keys(options).length > 0 ? 'Select insulation...' : 'Select heat source type first...';
      placeholder.disabled = true;
      placeholder.selected = true; // Default selection
      pipeInsulationTypeSelect.appendChild(placeholder);

      for (const key in options) {
          const option = document.createElement('option');
          option.value = key;
          let label = options[key].name;
          if (key === bestCaseKey) {
              label += ' (Best Case - Least Heat Transfer)';
          } else if (key === worstCaseKey) {
              label += ' (Worst Case - Most Heat Transfer)';
          }
          option.textContent = label;
          pipeInsulationTypeSelect.appendChild(option);
      }
      
      const noInsulationOption = document.createElement('option');
      noInsulationOption.value = 'none';
      noInsulationOption.textContent = 'None';
      pipeInsulationTypeSelect.appendChild(noInsulationOption);

      const customOption = document.createElement('option');
      customOption.value = 'other';
      customOption.textContent = 'Other (specify)...';
      pipeInsulationTypeSelect.appendChild(customOption);
      
      // FIX: Attempt to restore the previously selected value if it exists in the new set of options.
      if (Array.from(pipeInsulationTypeSelect.options).some(o => o.value === currentValue)) {
          pipeInsulationTypeSelect.value = currentValue;
      }
      
      handlePipeInsulationChange();
  };
  
  const handlePipeInsulationChange = () => {
    const selectedValue = pipeInsulationTypeSelect.value;
    const heatSourceType = document.querySelector<HTMLInputElement>('input[name="heatSourceType"]:checked')?.value;
    
    if (selectedValue === 'other') {
      customPipeInsulationWrapper?.classList.remove('hidden');
      insulationThermalConductivityDisplay?.classList.add('hidden');
      insulationThicknessWrapper?.classList.remove('hidden');
    } else if (selectedValue === 'none' || !selectedValue) {
      customPipeInsulationWrapper?.classList.add('hidden');
      insulationThermalConductivityDisplay?.classList.add('hidden');
      insulationThicknessWrapper?.classList.add('hidden');
    } else {
      customPipeInsulationWrapper?.classList.add('hidden');
      insulationThicknessWrapper?.classList.remove('hidden');
      
      let material;
      if (heatSourceType === 'steam' || heatSourceType === 'superHeatedHotWater') {
          material = insulationData.steam[selectedValue];
      } else if (heatSourceType === 'hotWater') {
          material = insulationData.hotWater[selectedValue];
      }
      
      if (material && insulationThermalConductivityDisplay) {
        insulationThermalConductivityDisplay.textContent = `Thermal Conductivity: ${material.thermalConductivity} BTU/hr·ft·°F`;
        insulationThermalConductivityDisplay.classList.remove('hidden');
      } else {
        insulationThermalConductivityDisplay?.classList.add('hidden');
      }
    }
  };
  pipeInsulationTypeSelect.addEventListener('change', handlePipeInsulationChange);

  // --- Heat Source Connection Types Logic ---
  const connectionTypesSelect = document.getElementById('connectionTypes') as HTMLSelectElement;
  const selectedConnectionTypesDisplay = document.getElementById('selectedConnectionTypesDisplay');
  const customConnectionTypesWrapper = document.getElementById('customConnectionTypesWrapper');

  const handleConnectionTypesChange = () => {
      if (!selectedConnectionTypesDisplay || !customConnectionTypesWrapper) return;
  
      selectedConnectionTypesDisplay.innerHTML = '';
      const selectedOptions = Array.from(connectionTypesSelect.selectedOptions);
      let isOtherSelected = false;
      let hasVisibleSelection = false;
  
      selectedOptions.forEach(option => {
          if (option.value === 'other') {
              isOtherSelected = true;
          } else if (option.value) {
              hasVisibleSelection = true;
              const tag = document.createElement('span');
              tag.className = 'selected-option-tag';
              tag.textContent = option.text;
  
              const removeBtn = document.createElement('button');
              removeBtn.className = 'remove-tag-btn';
              removeBtn.innerHTML = '&times;';
              removeBtn.type = 'button';
              removeBtn.setAttribute('aria-label', `Remove ${option.text}`);
              removeBtn.onclick = (e) => {
                  e.stopPropagation();
                  option.selected = false;
                  connectionTypesSelect.dispatchEvent(new Event('change'));
              };
  
              tag.appendChild(removeBtn);
              selectedConnectionTypesDisplay.appendChild(tag);
          }
      });
  
      if (hasVisibleSelection) {
          selectedConnectionTypesDisplay.classList.remove('hidden');
      } else {
          selectedConnectionTypesDisplay.classList.add('hidden');
      }
  
      if (isOtherSelected) {
          customConnectionTypesWrapper.classList.remove('hidden');
      } else {
          customConnectionTypesWrapper.classList.add('hidden');
      }
  };
  connectionTypesSelect.addEventListener('change', handleConnectionTypesChange);

  // --- Heat Loss Evidence Logic ---
  const heatLossSourceRadios = document.querySelectorAll<HTMLInputElement>('input[name="heatLossEvidenceSource"]');
  const customHeatLossSourceWrapper = document.getElementById('customHeatLossSourceWrapper');

  const handleHeatLossSourceChange = () => {
      const selectedRadio = document.querySelector<HTMLInputElement>('input[name="heatLossEvidenceSource"]:checked');
      if (customHeatLossSourceWrapper) {
          if (selectedRadio?.value === 'other') {
              customHeatLossSourceWrapper.classList.remove('hidden');
          } else {
              customHeatLossSourceWrapper.classList.add('hidden');
          }
      }
  };

  heatLossSourceRadios.forEach(radio => {
      radio.addEventListener('change', handleHeatLossSourceChange);
  });

  // --- Insulation Condition Source Logic ---
  const insulationConditionSourceRadios = document.querySelectorAll<HTMLInputElement>('input[name="insulationConditionSource"]');
  const customInsulationConditionSourceWrapper = document.getElementById('customInsulationConditionSourceWrapper');

  const handleInsulationConditionSourceChange = () => {
      const selectedRadio = document.querySelector<HTMLInputElement>('input[name="insulationConditionSource"]:checked');
      if (customInsulationConditionSourceWrapper) {
          if (selectedRadio?.value === 'other') {
              customInsulationConditionSourceWrapper.classList.remove('hidden');
          } else {
              customInsulationConditionSourceWrapper.classList.add('hidden');
          }
      }
  };
  insulationConditionSourceRadios.forEach(radio => {
      radio.addEventListener('change', handleInsulationConditionSourceChange);
  });

  // --- Trench Bedding Logic ---
  const createBeddingHandler = (
    typeRadioName: string,
    detailsWrapperId: string,
    customKRadioName: string,
    customKWrapperId: string,
    defaultKDisplayId: string
  ) => {
    const typeRadios = document.querySelectorAll<HTMLInputElement>(`input[name="${typeRadioName}"]`);
    const detailsWrapper = document.getElementById(detailsWrapperId);
    
    const handleTypeChange = () => {
      const selected = document.querySelector<HTMLInputElement>(`input[name="${typeRadioName}"]:checked`)?.value;
      detailsWrapper?.classList.toggle('hidden', selected !== 'sand');
    };
    
    typeRadios.forEach(radio => radio.addEventListener('change', handleTypeChange));
    
    const customKRadios = document.querySelectorAll<HTMLInputElement>(`input[name="${customKRadioName}"]`);
    const customKWrapper = document.getElementById(customKWrapperId);
    const defaultKDisplay = document.getElementById(defaultKDisplayId);

    const handleCustomKChange = () => {
        const selected = document.querySelector<HTMLInputElement>(`input[name="${customKRadioName}"]:checked`)?.value;
        customKWrapper?.classList.toggle('hidden', selected !== 'yes');
        defaultKDisplay?.classList.toggle('hidden', selected === 'yes');
    };

    customKRadios.forEach(radio => radio.addEventListener('change', handleCustomKChange));

    // Initial state
    handleTypeChange();
    handleCustomKChange();
  };

  createBeddingHandler(
    'heatSourceBeddingType', 'heatSourceBeddingDetailsWrapper',
    'heatSourceBeddingUseCustomK', 'customHeatSourceBeddingKWrapper', 'heatSourceBeddingDefaultKDisplay'
  );
  createBeddingHandler(
    'gasLineBeddingType', 'gasLineBeddingDetailsWrapper',
    'gasBeddingUseCustomK', 'customGasBeddingKWrapper', 'gasBeddingDefaultKDisplay'
  );


  // --- Soil Type Logic ---
  const createSoilTypeHandler = () => {
    const soilTypeSelect = document.getElementById('soilType') as HTMLSelectElement;
    const soilConductivityInput = document.getElementById('soilThermalConductivity') as HTMLInputElement;

    const updateConductivity = () => {
        if (!soilTypeSelect || !soilConductivityInput) return;
        const selectedOption = soilTypeSelect.options[soilTypeSelect.selectedIndex];
        // FIX: Add a check to ensure selectedOption is not null before getting attribute
        if (selectedOption) {
            const conductivity = selectedOption.getAttribute('data-conductivity');
            if (conductivity) {
                soilConductivityInput.value = conductivity;
            }
        }
    };

    soilTypeSelect.addEventListener('change', updateConductivity);
    updateConductivity(); // Initial call to set default value
  };
  createSoilTypeHandler();

  // --- Calculation Summary Logic ---
  const createSummaryItem = (label: string, value: any): HTMLElement | null => {
    let displayValue: string;
    
    if (value === null || value === undefined || value === '' || value === 'N/A') {
      displayValue = 'N/A';
    } else if (Array.isArray(value)) {
      displayValue = value.length > 0 ? value.join('\n') : 'N/A';
    } else {
      displayValue = String(value);
    }
    
    const item = document.createElement('div');
    item.className = 'summary-item';
  
    const labelEl = document.createElement('div');
    labelEl.className = 'summary-label';
    labelEl.textContent = label;
  
    const valueEl = document.createElement('div');
    valueEl.className = 'summary-value';
    valueEl.textContent = displayValue;
  
    item.appendChild(labelEl);
    item.appendChild(valueEl);
    return item;
  };
  
  const updateCalculationSummary = () => {
    const evaluationContainer = document.getElementById('summary-evaluation')!;
    const heatSourceContainer = document.getElementById('summary-heat-source')!;
    const gasLineContainer = document.getElementById('summary-gas-line')!;
    const soilContainer = document.getElementById('summary-soil')!;
    const fieldVisitContainer = document.getElementById('summary-field-visit')!;

    evaluationContainer.innerHTML = '';
    heatSourceContainer.innerHTML = '';
    gasLineContainer.innerHTML = '';
    soilContainer.innerHTML = '';
    fieldVisitContainer.innerHTML = '';
    
    const addItem = (container: HTMLElement, label: string, value: any) => {
        const item = createSummaryItem(label, value);
        if (item) container.appendChild(item);
    };

    const data = getFormData();
    const gasPipelineDiameterText = getSelectedText('gasPipelineDiameter');
    const heatSourcePipelineDiameterText = getSelectedText('pipelineDiameter');


    // Evaluation Information
    addItem(evaluationContainer, 'Date', data.date);
    addItem(evaluationContainer, `Evaluator Name${(data.evaluatorNames || []).length > 1 ? 's' : ''}`, data.evaluatorNames);
    addItem(evaluationContainer, 'Engineer’s Name', data.engineerName);
    addItem(evaluationContainer, 'Project Name', data.projectName);
    addItem(evaluationContainer, 'Project Location', data.projectLocation);
    addItem(evaluationContainer, 'Project Description', data.projectDescription);

    // Heat Source Data
    if (!data.isHeatSourceApplicable) {
        addItem(heatSourceContainer, 'Heat Source Status', 'No applicable heat source type selected.');
    } else {
        addItem(heatSourceContainer, 'Heat Source Type', data.heatSourceType);
        addItem(heatSourceContainer, 'Name of Individual(s) Providing Data', data.operatorName);
        addItem(heatSourceContainer, 'Operator Company Name', data.operatorCompanyName);
        addItem(heatSourceContainer, 'Operating Company Address', data.operatorCompanyAddress);
        addItem(heatSourceContainer, 'Contact Info (Phone or Email)', data.operatorContactInfo);
        const registered811Display = data.isRegistered811 !== 'N/A'
            ? `${data.isRegistered811} (${data.registered811Confirmation})`
            : 'N/A';
        addItem(heatSourceContainer, 'Registered Assets with 811 "DigSafe"', registered811Display);
        addItem(heatSourceContainer, 'Data Confirmation Date', data.confirmationDate);
        const maxTempDisplayValue = data.maxTemp ? `${data.maxTemp} (°F) (${data.tempType})` : 'N/A';
        addItem(heatSourceContainer, 'Max Operating Temp', maxTempDisplayValue);
        const maxPressureDisplayValue = data.maxPressure ? `${data.maxPressure} (psig) (${data.pressureType})` : 'N/A';
        addItem(heatSourceContainer, 'Max Operating Pressure', maxPressureDisplayValue);
        const heatSourceAgeDisplayValue = data.heatSourceAge ? `${data.heatSourceAge} (years) (${data.ageType})` : 'N/A';
        addItem(heatSourceContainer, 'Line Age', heatSourceAgeDisplayValue);
        const systemDutyCycleDisplayValue = data.systemDutyCycle ? `${data.systemDutyCycle} (${data.systemDutyCycleType})` : 'N/A';
        addItem(heatSourceContainer, 'System Duty Cycle / Uptime', systemDutyCycleDisplayValue);
        const pipeCasingInfoDisplayValue = data.pipeCasingInfo ? `${data.pipeCasingInfo} (${data.pipeCasingInfoType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Casing / Conduit Information', pipeCasingInfoDisplayValue);
        const heatLossEvidenceDisplayValue = data.heatLossEvidence ? `${data.heatLossEvidence} (Source: ${data.heatLossEvidenceSource})` : 'N/A';
        addItem(heatSourceContainer, 'Evidence of Surface Heat Loss', heatLossEvidenceDisplayValue);
        const pipelineDiameterDisplayValue = heatSourcePipelineDiameterText !== 'N/A' ? `${heatSourcePipelineDiameterText}" (${data.diameterType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipeline Nominal Diameter', pipelineDiameterDisplayValue);
        const heatSourcePipeMaterialDisplayValue = data.heatSourcePipeMaterial !== 'N/A' ? `${data.heatSourcePipeMaterial} (${data.materialType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Material', heatSourcePipeMaterialDisplayValue);
        const heatSourceWallThicknessDisplayValue = data.heatSourceWallThickness ? `${data.heatSourceWallThickness}" (${data.wallThicknessType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Wall Thickness', heatSourceWallThicknessDisplayValue);
        const connectionsValueText = data.connectionsValue.length > 0 ? data.connectionsValue.join('\n') : 'N/A';
        const connectionsDisplayValue = connectionsValueText !== 'N/A' ? `${connectionsValueText} (${data.connectionTypesType})` : 'N/A';
        addItem(heatSourceContainer, 'Line connection types', connectionsDisplayValue);
        const insulationTypeDisplayValue = data.insulationType !== 'N/A' ? `${data.insulationType} (${data.insulationTypeType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Insulation Type', insulationTypeDisplayValue);
        if (data.insulationType.startsWith('Other')) {
            addItem(heatSourceContainer, 'Custom Insulation Thermal Conductivity', data.customInsulationThermalConductivity);
        }
        if (data.insulationType !== 'None') {
            const insulationThicknessDisplayValue = data.insulationThickness ? `${data.insulationThickness}" (${data.insulationThicknessType})` : 'N/A';
            addItem(heatSourceContainer, 'Insulation Thickness', insulationThicknessDisplayValue);
            const insulationConditionDisplayValue = data.insulationCondition ? `${data.insulationCondition} (Source: ${data.insulationConditionSource})` : 'N/A';
            addItem(heatSourceContainer, 'Known Condition of Insulation', insulationConditionDisplayValue);
        }
        const heatSourceDepthDisplayValue = data.heatSourceDepth ? `${data.heatSourceDepth} (ft) (${data.depthType})` : 'N/A';
        addItem(heatSourceContainer, 'Depth to Pipe Centerline', heatSourceDepthDisplayValue);
        addItem(heatSourceContainer, 'Additional Information', data.additionalInfo);
    }

    // Gas Line Data
    const gasPipeOd = getGasPipeOd(gasPipelineDiameterText, data.gasPipeSizingStandard);
    const gasDiameterDisplay = gasPipelineDiameterText !== 'N/A'
      ? `${gasPipelineDiameterText}" (OD: ${!isNaN(gasPipeOd) ? gasPipeOd.toFixed(3) : 'N/A'} inches)`
      : 'N/A';

    addItem(gasLineContainer, 'Gas Line Operator Name', data.gasOperatorName);
    addItem(gasLineContainer, 'Max Operating Pressure (psig)', data.gasMaxPressure);
    addItem(gasLineContainer, 'Installation Year / Vintage', data.gasInstallationYear);
    addItem(gasLineContainer, 'Pipeline Nominal Diameter', gasDiameterDisplay);
    addItem(gasLineContainer, 'Pipe Sizing Standard', data.gasPipeSizingStandard.toUpperCase());
    addItem(gasLineContainer, 'Gas Pipeline Material', data.gasPipeMaterial);
    if (data.gasPipeWallThickness !== 'N/A') {
        addItem(gasLineContainer, 'Gas Pipe Wall Thickness (inches)', data.gasPipeWallThickness);
    }
    if (data.gasPipeSDR !== 'N/A') {
        addItem(gasLineContainer, 'Gas Pipe SDR', data.gasPipeSDR);
    }
    if (data.gasPipeMaterial.startsWith('Other')) {
        addItem(gasLineContainer, 'Custom Material Thermal Conductivity', data.customGasThermalConductivity || 'N/A');
    }

    const isPlastic = ['hdpe', 'mdpe', 'aldyl'].includes(data.gasPipeMaterialValue);
    if (isPlastic) {
        addItem(gasLineContainer, 'Material Continuous Temp Limit (°F)', data.gasPipeContinuousLimit);
        addItem(gasLineContainer, 'Common Utility Cap Temp (°F)', data.gasPipeUtilityCap);
        const operationalLimitItem = createSummaryItem('OPERATIONAL TEMP LIMIT (°F)', '70 (Max for all plastic pipes)');
        if (operationalLimitItem) {
            operationalLimitItem.style.backgroundColor = '#fcf8e3';
            operationalLimitItem.style.fontWeight = 'bold';
            gasLineContainer.appendChild(operationalLimitItem);
        }
        addItem(gasLineContainer, 'Key Notes', data.gasPipeNotes);
    }

    if (data.gasCoatingType !== 'N/A') {
        addItem(gasLineContainer, 'Coating Type', data.gasCoatingType);
        if (data.gasCoatingMaxTemp !== 'N/A' && data.gasCoatingMaxTemp !== 'Custom') {
            addItem(gasLineContainer, 'Coating Max Allowable Temp (°F)', data.gasCoatingMaxTemp);
        }
    }
    addItem(gasLineContainer, 'Orientation to Heat Source', data.gasLineOrientation);
    addItem(gasLineContainer, 'Depth to Pipe Centerline (feet)', data.depthOfBurialGasLine);
    if (data.parallelDistance !== 'N/A') {
        addItem(gasLineContainer, 'Centerline Distance (feet)', data.parallelDistance);
        addItem(gasLineContainer, 'Length of Parallel Section (feet)', data.parallelLength);
        if (data.parallelCoordinates && data.parallelCoordinates.length > 0) {
            const coordsText = data.parallelCoordinates
                .map(p => `${p.label}: ${p.lat || 'N/A'}, ${p.lng || 'N/A'}`)
                .join('\n');
            addItem(gasLineContainer, 'Parallel Section Coordinates', coordsText);
        }
    }
    if (data.latitude !== 'N/A' && data.longitude !== 'N/A') {
        const coordinates = (data.latitude && data.longitude) ? `${data.latitude}, ${data.longitude}` : 'N/A';
        addItem(gasLineContainer, 'Coordinates of Intersection (Lat, Lng)', coordinates);
    }
    if (data.lateralOffset !== 'N/A') {
        addItem(gasLineContainer, 'Lateral Offset (feet)', data.lateralOffset);
    }
    if (data.crossingAngle !== 'N/A') {
        addItem(gasLineContainer, 'Crossing Angle (degrees)', data.crossingAngle);
    }

    // Soil Data
    addItem(soilContainer, 'Native Soil Type Classification', getSelectedText('soilType').replace(/ \(.*\)/, ''));
    addItem(soilContainer, 'Native Soil Thermal Conductivity (BTU/hr·ft·°F)', data.soilThermalConductivity);
    addItem(soilContainer, 'Soil Moisture Content (%)', data.soilMoistureContent);
    addItem(soilContainer, 'Average Ground Temperature (°F)', data.averageGroundTemperature);
    addItem(soilContainer, 'Evidence of Water Infiltration/Frost Heave', data.waterInfiltration);
    if (data.waterInfiltrationComments) {
        addItem(soilContainer, 'Comments on Infiltration/Heave', data.waterInfiltrationComments);
    }
    
    // Heat Source Bedding
    addItem(soilContainer, 'Heat Source Trench Bedding', data.heatSourceBeddingType === 'sand' ? 'Installed with sand bedding' : 'Unknown');
    if (data.heatSourceBeddingType === 'sand') {
        const dims = `Above: ${data.heatSourceBeddingTop || 'N/A'} in, Below: ${data.heatSourceBeddingBottom || 'N/A'} in, Left: ${data.heatSourceBeddingLeft || 'N/A'} in, Right: ${data.heatSourceBeddingRight || 'N/A'} in`;
        addItem(soilContainer, 'Heat Source Bedding Dimensions', dims);
        const kValue = data.heatSourceBeddingUseCustomK === 'yes' ? `${data.heatSourceBeddingCustomK} (Custom)` : '0.20 (Default)';
        addItem(soilContainer, 'Heat Source Bedding Thermal Conductivity', kValue);
    }
    
    // Gas Line Bedding
    addItem(soilContainer, 'Gas Line Trench Bedding', data.gasLineBeddingType === 'sand' ? 'Installed with sand bedding' : 'Unknown');
    if (data.gasLineBeddingType === 'sand') {
        const dims = `Above: ${data.gasBeddingTop || 'N/A'} in, Below: ${data.gasBeddingBottom || 'N/A'} in, Left: ${data.gasBeddingLeft || 'N/A'} in, Right: ${data.gasBeddingRight || 'N/A'} in`;
        addItem(soilContainer, 'Gas Line Bedding Dimensions', dims);
        const kValue = data.gasBeddingUseCustomK === 'yes' ? `${data.gasBeddingCustomK} (Custom)` : '0.20 (Default)';
        addItem(soilContainer, 'Gas Line Bedding Thermal Conductivity', kValue);
    }

    // Field Visit Data
    addItem(fieldVisitContainer, 'Date of Visit', data.visitDate);
    addItem(fieldVisitContainer, 'Personnel on Site', data.sitePersonnel);
    addItem(fieldVisitContainer, 'Weather and Site Conditions', data.siteConditions);
    addItem(fieldVisitContainer, 'Field Observations and Notes', data.fieldObservations);
  };
  
  const updateAssessmentFormView = () => {
    const contentContainer = document.getElementById('assessment-form-content')!;
    contentContainer.innerHTML = ''; // Clear previous content

    // Helper functions to build the form document
    const createSection = (title: string, parent: HTMLElement) => {
        const section = document.createElement('div');
        section.className = 'form-section-printable';
        const header = document.createElement('h3');
        header.className = 'form-header-printable';
        header.textContent = title;
        section.appendChild(header);
        parent.appendChild(section);
        return section;
    };

    const createTextField = (label: string, parent: HTMLElement, notes = '') => {
        const field = document.createElement('div');
        field.className = 'form-field-printable text-field';
        field.innerHTML = `
            <label>${label}</label>
            <div class="input-line"></div>
            ${notes ? `<div class="field-notes">${notes}</div>` : ''}
        `;
        parent.appendChild(field);
    };
    
    const createTextAreaField = (label: string, parent: HTMLElement, rows = 4, notes = '') => {
        const field = document.createElement('div');
        field.className = 'form-field-printable text-area-field';
        let lines = '';
        for (let i = 0; i < rows; i++) {
            lines += '<div class="input-line"></div>';
        }
        field.innerHTML = `
            <label>${label}</label>
            <div class="input-box">${lines}</div>
            ${notes ? `<div class="field-notes">${notes}</div>` : ''}
        `;
        parent.appendChild(field);
    };

    const createChoiceField = (label: string, options: string[], parent: HTMLElement, notes = '', isMulti = false) => {
        const field = document.createElement('div');
        field.className = 'form-field-printable choice-field';
        let optionsHTML = options.map(opt => `
            <div class="choice-option">
                <span class="checkbox"></span>
                <span>${opt}</span>
            </div>
        `).join('');
        
        field.innerHTML = `
            <label>${label}</label>
            <div class="choices-container">${optionsHTML}</div>
            ${notes ? `<div class="field-notes">${notes}` : ''}
            ${isMulti ? `<br><em>(Check all that apply)</em></div>` : '</div>'}
        `;
        parent.appendChild(field);
    };

    // --- Build the form ---

    // Evaluation Information
    const evalSection = createSection('Evaluation Information', contentContainer);
    createTextField('Date:', evalSection);
    createTextField('Engineer’s Name:', evalSection);
    createTextField('Project Name:', evalSection);
    createTextField('Project Location:', evalSection);
    createTextAreaField('Project Description:', evalSection, 5);
    createTextAreaField('Evaluator(s) Name(s) & Affiliation(s):', evalSection, 3);
    
    // Heat Source Data
    const hsSection = createSection('Heat Source Data', contentContainer);
    createChoiceField('Buried Heat Source Type:', ['Steam Line', 'Hot Water Line', 'Super Heated Hot Water Line', 'Other: _______________'], hsSection);
    createTextField('Operator Company Name:', hsSection);
    createTextField('Operating Company Address:', hsSection);
    createTextField('Name of Individual(s) Providing Data:', hsSection);
    createTextField('Contact Info (Phone or Email):', hsSection);
    createChoiceField('Has operator Registered Assets with 811 "DigSafe"?', ['Yes', 'No', 'Unknown'], hsSection);
    createChoiceField('Confirmation of 811 Registration Status:', ['Confirmed with operator', 'Assumed', 'Unknown'], hsSection);
    createTextField('Data Confirmation Date:', hsSection);
    createTextField('Max Operating Temp (°F):', hsSection, 'Circle one: Unknown / Actual / Estimated');
    createTextField('Max Operating Pressure (psig):', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextField('Line Age (years):', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextAreaField('System Duty Cycle / Uptime:', hsSection, 3, 'e.g., "24/7", "Winter only". Circle one: Unknown / Actual / Assumed');
    createTextAreaField('Pipe Casing / Conduit Information:', hsSection, 3, 'Circle one: Unknown / Actual / Assumed');
    createTextAreaField('Evidence of Heat Loss at the Surface:', hsSection, 4, 'e.g., melted snow, stressed vegetation. Note source: Field Visit / Operator Info / Other');
    createTextField('Pipeline Nominal Diameter (inches):', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextField('Pipe Material:', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextField('Pipe Wall Thickness (inches):', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createChoiceField('Line connection types:', ['Butt-welded', 'Socket-welded', 'Flanged', 'Grooved Coupling', 'Brazed/Soldered', 'Plastic Fusion', 'Other: _________'], hsSection, 'Circle one: Unknown / Actual / Assumed', true);
    createTextField('Pipe Insulation Type:', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextField('Insulation Thickness (inches):', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextAreaField('Known Condition of Insulation:', hsSection, 3, 'e.g., degraded, water-logged. Note source: Unknown / Confirmed by Operator / Assumed / Other');
    createTextField('Depth from Ground Surface to Pipe Centerline (feet):', hsSection, 'Circle one: Unknown / Actual / Assumed');
    createTextAreaField('Additional Information Provided by the Operator:', hsSection, 4);

    // Gas Line Data
    const gasSection = createSection('Gas Line Data', contentContainer);
    createTextField('Gas Line Operator Name:', gasSection);
    createTextField('Gas Line Max Operating Pressure (psig):', gasSection);
    createTextField('Installation Year / Vintage:', gasSection);
    createTextField('Gas Line Pipeline Nominal Diameter (inches):', gasSection);
    createChoiceField('Pipe Sizing Standard:', ['IPS (Iron Pipe Size)', 'CTS (Copper Tube Size)'], gasSection);
    createTextField('Gas Pipeline Material:', gasSection);
    createTextField('Gas Pipe Wall Thickness (inches):', gasSection);
    createTextField('Gas Pipe SDR (if applicable):', gasSection);
    createTextField('Coating Type:', gasSection);
    createChoiceField('Gas Line Orientation to Heat Source:', ['Perpendicular', 'Parallel'], gasSection);
    createTextField('Depth of Gas Line to Pipe Centerline (feet):', gasSection);
    createTextField('If Parallel: Distance from Heat Source to Gas Line Centerline (feet):', gasSection);
    createTextField('If Parallel: Length of Parallel Section (feet):', gasSection);
    createTextAreaField('Coordinates of Intersection / Parallel Section:', gasSection, 3, 'Note Lat/Long for key points.');

    // Soil Data
    const soilSection = createSection('Soil & Bedding Data', contentContainer);
    createTextField('Native Soil Type Classification:', soilSection);
    createTextField('Native Soil Thermal Conductivity (BTU/hr·ft·°F):', soilSection);
    createTextField('Soil Moisture Content (%):', soilSection);
    createTextField('Average Ground Temperature (°F):', soilSection);
    createTextAreaField('Evidence of water infiltration or frost heave:', soilSection, 3, 'Circle one: Yes / No / Unknown');
    createTextAreaField('Heat Source Trench Bedding:', soilSection, 2, 'Describe bedding type (e.g., sand) and dimensions if known.');
    createTextAreaField('Gas Line Trench Bedding:', soilSection, 2, 'Describe bedding type (e.g., sand) and dimensions if known.');

    // Field Visit
    const fieldSection = createSection('Field Visit', contentContainer);
    createTextField('Date of Visit:', fieldSection);
    createTextAreaField('Personnel on Site:', fieldSection, 3);
    createTextAreaField('Weather and Site Conditions:', fieldSection, 4);
    createTextAreaField('Field Observations and Notes:', fieldSection, 10, 'Note other utilities, surface conditions, markers, visual indicators, measurements, etc.');

    // Photo Log section
    const photoSection = createSection('Field Photos Log / Sketches', contentContainer);
    let photoLogHTML = '<div class="photo-log-grid">';
    for (let i = 1; i <= 10; i++) {
        photoLogHTML += `
            <div class="photo-log-entry">
                <div class="photo-box">Photo #${i} / Sketch</div>
                <div class="photo-desc-area">
                    <label>Description:</label>
                    <div class="input-line"></div>
                    <div class="input-line"></div>
                    <div class="input-line"></div>
                </div>
            </div>
        `;
    }
    photoLogHTML += '</div>';
    photoSection.innerHTML += photoLogHTML;
  };
  
  document.getElementById('exportPdfButton')?.addEventListener('click', () => {
    const element = document.getElementById('assessment-form-content');
    const button = document.getElementById('exportPdfButton') as HTMLButtonElement;
    
    if (!element || !button) return;

    button.textContent = 'Exporting...';
    button.disabled = true;

    const opt = {
      margin:       [0.5, 0.5, 0.5, 0.5],
      filename:     `assessment-form-${new Date().toISOString().split('T')[0]}.pdf`,
      image:        { type: 'jpeg', quality: 0.98 },
      html2canvas:  { scale: 2, useCORS: true, letterRendering: true },
      jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' },
      pagebreak:    { mode: ['avoid-all', 'css', 'legacy'] }
    };
    
    // The html2pdf library is loaded from a CDN and will be available on the window object.
    (window as any).html2pdf().from(element).set(opt).save().then(() => {
      button.textContent = 'Export to PDF';
      button.disabled = false;
    }).catch((err: Error) => {
        console.error("Failed to export PDF:", err);
        alert("An error occurred during PDF export. Please check the console for details.");
        button.textContent = 'Export to PDF';
        button.disabled = false;
    });
  });

  // --- Admin Login Logic ---
  const adminExecuteButton = document.getElementById('adminExecuteButton');
  const adminLogoutButton = document.getElementById('adminLogoutButton');
  const adminPasscodeInput = document.getElementById('adminPasscode') as HTMLInputElement;
  const adminLoginError = document.getElementById('adminLoginError');
  const adminLoginWrapper = document.getElementById('admin-login-wrapper');
  const adminContentWrapper = document.getElementById('admin-content-wrapper');
  const apiKeyInput = document.getElementById('apiKeyInput') as HTMLInputElement;
  const saveApiKeyButton = document.getElementById('saveApiKeyButton');
  const apiKeyStatus = document.getElementById('apiKeyStatus');

  const updateApiKeyStatus = () => {
    if (!apiKeyStatus) return;
    const key = sessionStorage.getItem('gemini-api-key');
    if (ai && key) {
        const maskedKey = `${key.substring(0, 4)}...${key.substring(key.length - 4)}`;
        apiKeyStatus.innerHTML = `API Key is set for this session: <code>${maskedKey}</code>`;
        apiKeyStatus.className = 'api-key-status set';
    } else {
        apiKeyStatus.textContent = 'API Key is not set. AI features are disabled.';
        apiKeyStatus.className = 'api-key-status not-set';
        ai = null; // Ensure AI is null if no key
    }
  };

  const initializeAiClient = (apiKey: string): boolean => {
    if (!apiKey) {
        ai = null;
        updateApiKeyStatus();
        return false;
    }
    try {
        ai = new GoogleGenAI({ apiKey });
        return true;
    } catch (error) {
        console.error("Failed to initialize GoogleGenAI:", error);
        alert(`Failed to initialize AI client. Check the console for details. Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
        ai = null;
        sessionStorage.removeItem('gemini-api-key');
        updateApiKeyStatus();
        return false;
    }
  };

  const handleSaveApiKey = () => {
    const apiKey = apiKeyInput.value.trim();
    if (!apiKey) {
      alert("Please enter a valid API key.");
      return;
    }
    if (initializeAiClient(apiKey)) {
        sessionStorage.setItem('gemini-api-key', apiKey);
        alert("API Key saved successfully for this session.");
        enableAdminFeatures();
    } else {
        sessionStorage.removeItem('gemini-api-key');
        disableAdminFeatures();
    }
    apiKeyInput.value = ''; // Clear input for security
    updateApiKeyStatus();
  };

  const enableAdminFeatures = () => {
    document.querySelectorAll('.reword-button').forEach(btn => {
        btn.classList.remove('hidden');
    });
    handleFinalReportTabClick(); // Update final report tab state
  };
  
  const disableAdminFeatures = () => {
      document.querySelectorAll('.reword-button').forEach(btn => {
          btn.classList.add('hidden');
      });
      ai = null;
      handleFinalReportTabClick();
  };

  const handleAdminExecute = () => {
    if (adminPasscodeInput.value === '0665') {
        isAdminAuthenticated = true;
        adminLoginWrapper?.classList.add('hidden');
        adminContentWrapper?.classList.remove('hidden');
        if (adminLoginError) adminLoginError.classList.add('hidden');
        adminPasscodeInput.value = '';
        
        // On login, check if a key is already in the session
        const existingKey = sessionStorage.getItem('gemini-api-key');
        if (existingKey) {
            if (initializeAiClient(existingKey)) {
                enableAdminFeatures();
            } else {
                disableAdminFeatures();
            }
        } else {
          disableAdminFeatures();
        }
        updateApiKeyStatus();

    } else if (adminLoginError) {
        adminLoginError.textContent = 'Incorrect passcode.';
        adminLoginError.classList.remove('hidden');
    }
  };

  const handleAdminLogout = () => {
    isAdminAuthenticated = false;
    ai = null;
    sessionStorage.removeItem('gemini-api-key');
    adminLoginWrapper?.classList.remove('hidden');
    adminContentWrapper?.classList.add('hidden');
    disableAdminFeatures();
    updateApiKeyStatus();
  };
  
  // Add event listeners for admin actions
  adminExecuteButton?.addEventListener('click', handleAdminExecute);
  adminLogoutButton?.addEventListener('click', handleAdminLogout);
  saveApiKeyButton?.addEventListener('click', handleSaveApiKey);
  
  // No AI client initialization on page load, must be done by admin.
  updateApiKeyStatus();

  // Initial UI setup calls
  updateEvaluatorInputs();
  toggleConditionalFields();
  handleGasOperatorNameChange();
  handleGasPipeMaterialChange();
  handleGasCoatingTypeChange();
  handleGasSdrChange();
  handleGasLineOrientationChange();
  updateParallelCoordinateInputs();
  handleConnectionTypesChange();
  handleHeatLossSourceChange();
  handleInsulationConditionSourceChange();

  // Activate the first tab by default
  if (tabButtons.length > 0 && tabPanels.length > 0) {
    tabButtons[0].classList.add('active');
    tabPanels[0].classList.add('active');
  }
  updateAssessmentFormView(); // Populate the printable form view initially.
});