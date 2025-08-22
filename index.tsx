/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/
import { GoogleGenAI, Type } from "@google/genai";

// Moved from inside DOMContentLoaded to be globally available.
const getFormData = () => {
  const getEl = (id: string) => document.getElementById(id);
  const getVal = (id: string) => (getEl(id) as HTMLInputElement)?.value || '';
  const getSelectedText = (id: string) => {
      const select = getEl(id) as HTMLSelectElement;
      return select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : 'N/A';
  };
  const getRadioVal = (name: string) => {
      const checked = document.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`);
      return checked ? checked.value : '';
  };
  const getRadioLabel = (name: string) => {
      const checked = document.querySelector<HTMLInputElement>(`input[name="${name}"]:checked`);
      return checked ? (document.querySelector(`label[for="${checked.id}"]`)?.textContent || 'N/A') : 'N/A';
  };

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
      heatSourceWallThickness = `Other: ${getVal('customWallThickness')}`;
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
  if (getRadioVal('insulationConditionSource') === 'other') {
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
      pipelineDiameter: getSelectedText('pipelineDiameter'),
      materialType: getRadioLabel('materialType'),
      heatSourcePipeMaterial,
      wallThicknessType: getRadioLabel('wallThicknessType'),
      heatSourceWallThickness,
      connectionTypesType: getRadioLabel('connectionTypesType'),
      connectionsValue,
      insulationTypeType: getRadioLabel('insulationTypeType'),
      insulationType,
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
      gasDoc: getSelectedText('gasDoc'),
      gasMaxPressure: getVal('gasMaxPressure'),
      gasInstallationYear: getVal('gasInstallationYear'),
      gasPipelineDiameter: getSelectedText('gasPipelineDiameter'),
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
      // Soil
      soilType: getSelectedText('soilType').replace(/ \(.*/, ''),
      soilThermalConductivity: getVal('soilThermalConductivity'),
      soilMoistureContent: getVal('soilMoistureContent'),
      groundSurfaceTemperature: getVal('groundSurfaceTemperature'),
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

// Moved from inside DOMContentLoaded to be globally available.
interface ScenarioDetails {
  Q: number;
  T_gas_line: number;
  delta_T: number;
  R_pipe_wall: number;
  R_insulation: number;
  R_soil_hs: number;
  R_total: number;
  D_hs_outer_ft: number;
  D_hs_inner_ft: number;
  D_ins_outer_ft?: number;
  R_surface_ft: number;
  x_target_edge: number;
  y_target_edge: number;
  r_equiv_ft?: number;
  R_bedding?: number;
  R_soil_from_bedding?: number;
}

const createResizableImage = (src: string, container: HTMLElement) => {
  const wrapper = document.createElement('div');
  wrapper.className = 'resizable-image-wrapper';

  const img = document.createElement('img');
  img.src = src;
  img.alt = 'Field visit photo preview';
  
  const handle = document.createElement('div');
  handle.className = 'resize-handle';
  
  wrapper.appendChild(img);
  wrapper.appendChild(handle);
  container.appendChild(wrapper);

  // --- Resizing Logic ---
  const startResize = (e: MouseEvent) => {
      e.preventDefault();
      
      let startX = e.clientX;
      let startY = e.clientY;
      let startWidth = wrapper.offsetWidth;
      let startHeight = wrapper.offsetHeight;

      const doResize = (moveEvent: MouseEvent) => {
          const newWidth = startWidth + (moveEvent.clientX - startX);
          const newHeight = startHeight + (moveEvent.clientY - startY);
          wrapper.style.width = `${newWidth}px`;
          wrapper.style.height = `${newHeight}px`;
      };
      
      const stopResize = () => {
          document.documentElement.removeEventListener('mousemove', doResize, false);
          document.documentElement.removeEventListener('mouseup', stopResize, false);
          document.body.classList.remove('resizing');
      };

      document.documentElement.addEventListener('mousemove', doResize, false);
      document.documentElement.addEventListener('mouseup', stopResize, false);
      document.body.classList.add('resizing');
  };

  handle.addEventListener('mousedown', startResize, false);
};

document.addEventListener('DOMContentLoaded', () => {
  let isAdminAuthenticated = false;
  let targetTextareaIdForReword: string | null = null;
  let recommendationCounter = 0;

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
    if (maxTempLabel) maxTempLabel.textContent = `${sourceName} Max Operating Temp (in Deg F) confirmed by the operator:`;
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
      let options: { [key: string]: { name: string; thermalConductivity: number } } = {};
      if (heatSourceType === 'steam' || heatSourceType === 'superHeatedHotWater') {
          options = insulationData.steam;
      } else if (heatSourceType === 'hotWater') {
          options = insulationData.hotWater;
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
          
          // In case all values are the same, don't label anything
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
      placeholder.selected = true;
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
      
      // Reset dependent fields
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
        const conductivity = selectedOption.getAttribute('data-conductivity');
        if (conductivity) {
            soilConductivityInput.value = conductivity;
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
        const maxTempDisplayValue = data.maxTemp ? `${data.maxTemp} (${data.tempType})` : 'N/A';
        addItem(heatSourceContainer, 'Max Operating Temp (°F)', maxTempDisplayValue);
        const maxPressureDisplayValue = data.maxPressure ? `${data.maxPressure} (${data.pressureType})` : 'N/A';
        addItem(heatSourceContainer, 'Max Operating Pressure (psig)', maxPressureDisplayValue);
        const heatSourceAgeDisplayValue = data.heatSourceAge ? `${data.heatSourceAge} (${data.ageType})` : 'N/A';
        addItem(heatSourceContainer, 'Line Age (years)', heatSourceAgeDisplayValue);
        const systemDutyCycleDisplayValue = data.systemDutyCycle ? `${data.systemDutyCycle} (${data.systemDutyCycleType})` : 'N/A';
        addItem(heatSourceContainer, 'System Duty Cycle / Uptime', systemDutyCycleDisplayValue);
        const pipeCasingInfoDisplayValue = data.pipeCasingInfo ? `${data.pipeCasingInfo} (${data.pipeCasingInfoType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Casing / Conduit Information', pipeCasingInfoDisplayValue);
        const heatLossEvidenceDisplayValue = data.heatLossEvidence ? `${data.heatLossEvidence} (Source: ${data.heatLossEvidenceSource})` : 'N/A';
        addItem(heatSourceContainer, 'Evidence of Surface Heat Loss', heatLossEvidenceDisplayValue);
        const pipelineDiameterDisplayValue = data.pipelineDiameter !== 'N/A' ? `${data.pipelineDiameter} (${data.diameterType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipeline Nominal Diameter (inches)', pipelineDiameterDisplayValue);
        const heatSourcePipeMaterialDisplayValue = data.heatSourcePipeMaterial !== 'N/A' ? `${data.heatSourcePipeMaterial} (${data.materialType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Material', heatSourcePipeMaterialDisplayValue);
        const heatSourceWallThicknessDisplayValue = data.heatSourceWallThickness ? `${data.heatSourceWallThickness} (${data.wallThicknessType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Wall Thickness (inches)', heatSourceWallThicknessDisplayValue);
        const connectionsValueText = data.connectionsValue.length > 0 ? data.connectionsValue.join('\n') : 'N/A';
        const connectionsDisplayValue = connectionsValueText !== 'N/A' ? `${connectionsValueText} (${data.connectionTypesType})` : 'N/A';
        addItem(heatSourceContainer, 'Line connection types', connectionsDisplayValue);
        const insulationTypeDisplayValue = data.insulationType !== 'N/A' ? `${data.insulationType} (${data.insulationTypeType})` : 'N/A';
        addItem(heatSourceContainer, 'Pipe Insulation Type', insulationTypeDisplayValue);
        if (data.insulationType.startsWith('Other')) {
            addItem(heatSourceContainer, 'Custom Insulation Thermal Conductivity', data.customInsulationThermalConductivity);
        }
        if (data.insulationType !== 'None') {
            const insulationThicknessDisplayValue = data.insulationThickness ? `${data.insulationThickness} (${data.insulationThicknessType})` : 'N/A';
            addItem(heatSourceContainer, 'Insulation Thickness (inches)', insulationThicknessDisplayValue);
            const insulationConditionDisplayValue = data.insulationCondition ? `${data.insulationCondition} (Source: ${data.insulationConditionSource})` : 'N/A';
            addItem(heatSourceContainer, 'Known Condition of Insulation', insulationConditionDisplayValue);
        }
        const heatSourceDepthDisplayValue = data.heatSourceDepth ? `${data.heatSourceDepth} (${data.depthType})` : 'N/A';
        addItem(heatSourceContainer, 'Depth from Ground Surface to Pipe Centerline (feet)', heatSourceDepthDisplayValue);
        addItem(heatSourceContainer, 'Additional Information', data.additionalInfo);
    }

    // Gas Line Data
    const gasPipeOd = getGasPipeOd(data.gasPipelineDiameter, data.gasPipeSizingStandard);
    const gasDiameterDisplay = data.gasPipelineDiameter !== 'N/A'
      ? `${data.gasPipelineDiameter} (OD: ${!isNaN(gasPipeOd) ? gasPipeOd.toFixed(3) : 'N/A'} in)`
      : 'N/A';

    addItem(gasLineContainer, 'Gas Line Operator Name', data.gasOperatorName);
    addItem(gasLineContainer, 'Distribution Operating Company (DOC)', data.gasDoc);
    addItem(gasLineContainer, 'Gas Line Max Operating Pressure (psig)', data.gasMaxPressure);
    addItem(gasLineContainer, 'Installation Year / Vintage', data.gasInstallationYear);
    addItem(gasLineContainer, 'Gas Line Pipeline Nominal Diameter (in)', gasDiameterDisplay);
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
    addItem(gasLineContainer, 'Gas Line Orientation to Heat Source', data.gasLineOrientation);
    addItem(gasLineContainer, 'Depth of Gas Line to Pipe Centerline (feet)', data.depthOfBurialGasLine);
    if (data.parallelDistance !== 'N/A') {
        addItem(gasLineContainer, 'Distance from Heat Source to Gas Line Centerline (feet)', data.parallelDistance);
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

    // Soil Data
    addItem(soilContainer, 'Native Soil Type Classification', data.soilType);
    addItem(soilContainer, 'Native Soil Thermal Conductivity (BTU/hr·ft·°F)', data.soilThermalConductivity);
    addItem(soilContainer, 'Soil Moisture Content (%)', data.soilMoistureContent);
    addItem(soilContainer, 'Average Ground Surface Temperature (°F)', data.groundSurfaceTemperature);
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
    contentContainer.innerHTML = '';
  
    const data = getFormData();
  
    const createSection = (title: string) => {
      const sectionDiv = document.createElement('div');
      sectionDiv.className = 'summary-section';
      const titleEl = document.createElement('h3');
      titleEl.textContent = title;
      sectionDiv.appendChild(titleEl);
      contentContainer.appendChild(sectionDiv);
      return sectionDiv;
    };
  
    const addItem = (container: HTMLElement, label: string, value: any) => {
      const item = createSummaryItem(label, value);
      if (item) container.appendChild(item);
    };
  
    // --- Evaluation Section ---
    const evalSection = createSection('Evaluation Information');
    addItem(evalSection, 'Date', data.date);
    addItem(evalSection, `Evaluator Name${(data.evaluatorNames || []).length > 1 ? 's' : ''}`, data.evaluatorNames);
    addItem(evalSection, 'Engineer’s Name', data.engineerName);
    addItem(evalSection, 'Project Name', data.projectName);
    addItem(evalSection, 'Project Location', data.projectLocation);
    addItem(evalSection, 'Project Description', data.projectDescription);
  
    // --- Heat Source Section ---
    const hsSection = createSection('Heat Source Data');
    if (!data.isHeatSourceApplicable) {
      addItem(hsSection, 'Heat Source Status', 'No applicable heat source type selected.');
    } else {
      addItem(hsSection, 'Heat Source Type', data.heatSourceType);
      addItem(hsSection, 'Name of Individual(s) Providing Data', data.operatorName);
      addItem(hsSection, 'Operator Company Name', data.operatorCompanyName);
      addItem(hsSection, 'Operating Company Address', data.operatorCompanyAddress);
      addItem(hsSection, 'Contact Info (Phone or Email)', data.operatorContactInfo);
      const registered811Display = data.isRegistered811 !== 'N/A'
            ? `${data.isRegistered811} (${data.registered811Confirmation})`
            : 'N/A';
      addItem(hsSection, 'Has operator Registered Assets with 811 "DigSafe"?', registered811Display);
      addItem(hsSection, 'Data Confirmation Date', data.confirmationDate);
      const maxTempDisplayValue = data.maxTemp ? `${data.maxTemp} (${data.tempType})` : 'N/A';
      addItem(hsSection, 'Max Operating Temp (°F)', maxTempDisplayValue);
      const maxPressureDisplayValue = data.maxPressure ? `${data.maxPressure} (${data.pressureType})` : 'N/A';
      addItem(hsSection, 'Max Operating Pressure (psig)', maxPressureDisplayValue);
      const heatSourceAgeDisplayValue = data.heatSourceAge ? `${data.heatSourceAge} (${data.ageType})` : 'N/A';
      addItem(hsSection, 'Line Age (years)', heatSourceAgeDisplayValue);
      const systemDutyCycleDisplayValue = data.systemDutyCycle ? `${data.systemDutyCycle} (${data.systemDutyCycleType})` : 'N/A';
      addItem(hsSection, 'System Duty Cycle / Uptime', systemDutyCycleDisplayValue);
      const pipeCasingInfoDisplayValue = data.pipeCasingInfo ? `${data.pipeCasingInfo} (${data.pipeCasingInfoType})` : 'N/A';
      addItem(hsSection, 'Pipe Casing / Conduit Information', pipeCasingInfoDisplayValue);
      const heatLossEvidenceDisplayValue = data.heatLossEvidence ? `${data.heatLossEvidence} (Source: ${data.heatLossEvidenceSource})` : 'N/A';
      addItem(hsSection, 'Evidence of Surface Heat Loss', heatLossEvidenceDisplayValue);
      const pipelineDiameterDisplayValue = data.pipelineDiameter !== 'N/A' ? `${data.pipelineDiameter} (${data.diameterType})` : 'N/A';
      addItem(hsSection, 'Pipeline Nominal Diameter (inches)', pipelineDiameterDisplayValue);
      const heatSourcePipeMaterialDisplayValue = data.heatSourcePipeMaterial !== 'N/A' ? `${data.heatSourcePipeMaterial} (${data.materialType})` : 'N/A';
      addItem(hsSection, 'Pipe Material', heatSourcePipeMaterialDisplayValue);
      const heatSourceWallThicknessDisplayValue = data.heatSourceWallThickness ? `${data.heatSourceWallThickness} (${data.wallThicknessType})` : 'N/A';
      addItem(hsSection, 'Pipe Wall Thickness (inches)', heatSourceWallThicknessDisplayValue);
      const connectionsValueText = data.connectionsValue.length > 0 ? data.connectionsValue.join('\n') : 'N/A';
      const connectionsDisplayValue = connectionsValueText !== 'N/A' ? `${connectionsValueText} (${data.connectionTypesType})` : 'N/A';
      addItem(hsSection, 'Line connection types', connectionsDisplayValue);
      const insulationTypeDisplayValue = data.insulationType !== 'N/A' ? `${data.insulationType} (${data.insulationTypeType})` : 'N/A';
      addItem(hsSection, 'Pipe Insulation Type', insulationTypeDisplayValue);
      if (data.insulationType.startsWith('Other')) {
          addItem(hsSection, 'Custom Insulation Thermal Conductivity', data.customInsulationThermalConductivity);
      }
      if (data.insulationType !== 'None') {
          const insulationThicknessDisplayValue = data.insulationThickness ? `${data.insulationThickness} (${data.insulationThicknessType})` : 'N/A';
          addItem(hsSection, 'Insulation Thickness (inches)', insulationThicknessDisplayValue);
          const insulationConditionDisplayValue = data.insulationCondition ? `${data.insulationCondition} (Source: ${data.insulationConditionSource})` : 'N/A';
          addItem(hsSection, 'Known Condition of Insulation', insulationConditionDisplayValue);
      }
      const heatSourceDepthDisplayValue = data.heatSourceDepth ? `${data.heatSourceDepth} (${data.depthType})` : 'N/A';
      addItem(hsSection, 'Depth from Ground Surface to Pipe Centerline (feet)', heatSourceDepthDisplayValue);
      addItem(hsSection, 'Additional Information', data.additionalInfo);
    }
  
    // --- Gas Line Section ---
    const gasSection = createSection('Gas Line Data');
    const gasPipeOd = getGasPipeOd(data.gasPipelineDiameter, data.gasPipeSizingStandard);
    const gasDiameterDisplay = data.gasPipelineDiameter !== 'N/A'
      ? `${data.gasPipelineDiameter} (OD: ${!isNaN(gasPipeOd) ? gasPipeOd.toFixed(3) : 'N/A'} in)`
      : 'N/A';
    addItem(gasSection, 'Gas Line Operator Name', data.gasOperatorName);
    addItem(gasSection, 'Distribution Operating Company (DOC)', data.gasDoc);
    addItem(gasSection, 'Gas Line Max Operating Pressure (psig)', data.gasMaxPressure);
    addItem(gasSection, 'Installation Year / Vintage', data.gasInstallationYear);
    addItem(gasSection, 'Gas Line Pipeline Nominal Diameter (in)', gasDiameterDisplay);
    addItem(gasSection, 'Pipe Sizing Standard', data.gasPipeSizingStandard.toUpperCase());
    addItem(gasSection, 'Gas Pipeline Material', data.gasPipeMaterial);
    if (data.gasPipeWallThickness !== 'N/A') {
        addItem(gasSection, 'Gas Pipe Wall Thickness (inches)', data.gasPipeWallThickness);
    }
    if (data.gasPipeSDR !== 'N/A') {
        addItem(gasSection, 'Gas Pipe SDR', data.gasPipeSDR);
    }
    if (data.gasPipeMaterial.startsWith('Other')) {
        addItem(gasSection, 'Custom Material Thermal Conductivity', data.customGasThermalConductivity || 'N/A');
    }
    
    const isPlastic = ['hdpe', 'mdpe', 'aldyl'].includes(data.gasPipeMaterialValue);
    if (isPlastic) {
        addItem(gasSection, 'Material Continuous Temp Limit (°F)', data.gasPipeContinuousLimit);
        addItem(gasSection, 'Common Utility Cap Temp (°F)', data.gasPipeUtilityCap);
        const operationalLimitItem = createSummaryItem('OPERATIONAL TEMP LIMIT (°F)', '70 (Max for all plastic pipes)');
        if (operationalLimitItem) {
            operationalLimitItem.style.backgroundColor = '#fcf8e3';
            operationalLimitItem.style.fontWeight = 'bold';
            gasSection.appendChild(operationalLimitItem);
        }
        addItem(gasSection, 'Key Notes', data.gasPipeNotes);
    }
    
    if (data.gasCoatingType !== 'N/A') {
        addItem(gasSection, 'Coating Type', data.gasCoatingType);
        if (data.gasCoatingMaxTemp !== 'N/A' && data.gasCoatingMaxTemp !== 'Custom') {
            addItem(gasSection, 'Coating Max Allowable Temp (°F)', data.gasCoatingMaxTemp);
        }
    }
    addItem(gasSection, 'Gas Line Orientation to Heat Source', data.gasLineOrientation);
    addItem(gasSection, 'Depth of Gas Line to Pipe Centerline (feet)', data.depthOfBurialGasLine);
    if (data.parallelDistance !== 'N/A') {
        addItem(gasSection, 'Distance from Heat Source to Gas Line Centerline (feet)', data.parallelDistance);
        addItem(gasSection, 'Length of Parallel Section (feet)', data.parallelLength);
        if (data.parallelCoordinates && data.parallelCoordinates.length > 0) {
            const coordsText = data.parallelCoordinates
                .map(p => `${p.label}: ${p.lat || 'N/A'}, ${p.lng || 'N/A'}`)
                .join('\n');
            addItem(gasSection, 'Parallel Section Coordinates', coordsText);
        }
    }
    if (data.latitude !== 'N/A' && data.longitude !== 'N/A') {
        const coordinates = (data.latitude && data.longitude) ? `${data.latitude}, ${data.longitude}` : 'N/A';
        addItem(gasSection, 'Coordinates of Intersection (Lat, Lng)', coordinates);
    }
  
    // --- Soil Section ---
    const soilSection = createSection('Soil & Bedding Data');
    addItem(soilSection, 'Native Soil Type Classification', data.soilType);
    addItem(soilSection, 'Native Soil Thermal Conductivity (BTU/hr·ft·°F)', data.soilThermalConductivity);
    addItem(soilSection, 'Soil Moisture Content (%)', data.soilMoistureContent);
    addItem(soilSection, 'Average Ground Surface Temperature (°F)', data.groundSurfaceTemperature);
    addItem(soilSection, 'Evidence of Water Infiltration/Frost Heave', data.waterInfiltration);
    if (data.waterInfiltrationComments) {
        addItem(soilSection, 'Comments on Infiltration/Heave', data.waterInfiltrationComments);
    }
    addItem(soilSection, 'Heat Source Trench Bedding', data.heatSourceBeddingType === 'sand' ? 'Installed with sand bedding' : 'Unknown');
    if (data.heatSourceBeddingType === 'sand') {
        const dims = `Above: ${data.heatSourceBeddingTop || 'N/A'} in, Below: ${data.heatSourceBeddingBottom || 'N/A'} in, Left: ${data.heatSourceBeddingLeft || 'N/A'} in, Right: ${data.heatSourceBeddingRight || 'N/A'} in`;
        addItem(soilSection, 'Heat Source Bedding Dimensions', dims);
        const kValue = data.heatSourceBeddingUseCustomK === 'yes' ? `${data.heatSourceBeddingCustomK} (Custom)` : '0.20 (Default)';
        addItem(soilSection, 'Heat Source Bedding Thermal Conductivity', kValue);
    }
    addItem(soilSection, 'Gas Line Trench Bedding', data.gasLineBeddingType === 'sand' ? 'Installed with sand bedding' : 'Unknown');
    if (data.gasLineBeddingType === 'sand') {
        const dims = `Above: ${data.gasBeddingTop || 'N/A'} in, Below: ${data.gasBeddingBottom || 'N/A'} in, Left: ${data.gasBeddingLeft || 'N/A'} in, Right: ${data.gasBeddingRight || 'N/A'} in`;
        addItem(soilSection, 'Gas Line Bedding Dimensions', dims);
        const kValue = data.gasBeddingUseCustomK === 'yes' ? `${data.gasBeddingCustomK} (Custom)` : '0.20 (Default)';
        addItem(soilSection, 'Gas Line Bedding Thermal Conductivity', kValue);
    }
  
    // --- Field Visit Section ---
    const fieldSection = createSection('Field Visit');
    addItem(fieldSection, 'Date of Visit', data.visitDate);
    addItem(fieldSection, 'Personnel on Site', data.sitePersonnel);
    addItem(fieldSection, 'Weather and Site Conditions', data.siteConditions);
    addItem(fieldSection, 'Field Observations and Notes', data.fieldObservations);
  
    // --- Recommendations Section ---
    const recSection = createSection('Recommendations');
    if (data.recommendations && data.recommendations.length > 0 && data.recommendations.some(r => r.trim() !== '')) {
      data.recommendations.forEach((rec: string, index: number) => {
        if (rec.trim() !== '') {
          addItem(recSection, `Recommendation #${index + 1}`, rec);
        }
      });
    } else {
      addItem(recSection, 'Recommendations', 'No specific recommendations provided.');
    }
  };
  
  document.getElementById('printAssessmentFormButton')?.addEventListener('click', () => {
    document.body.classList.add('printing-assessment-form');
    window.print();
    document.body.classList.remove('printing-assessment-form');
  });

  // --- Admin Login Logic ---
  const adminLoginButton = document.getElementById('adminLoginButton');
  const adminPasscodeInput = document.getElementById('adminPasscode') as HTMLInputElement;
  const adminLoginError = document.getElementById('adminLoginError');
  const adminLoginWrapper = document.getElementById('admin-login-wrapper');
  const adminContentWrapper = document.getElementById('admin-content-wrapper');

  const enableAdminFeatures = () => {
    document.querySelectorAll('.reword-button').forEach(btn => {
        btn.classList.remove('hidden');
    });
    handleFinalReportTabClick(); // Update final report tab state
  };

  const handleAdminLogin = () => {
    if (adminPasscodeInput.value === '0665') {
      isAdminAuthenticated = true;
      adminLoginWrapper?.classList.add('hidden');
      adminContentWrapper?.classList.remove('hidden');
      adminLoginError?.classList.add('hidden');
      adminPasscodeInput.value = ''; // Clear password on success
      enableAdminFeatures();
    } else {
      if (adminLoginError) {
        adminLoginError.textContent = 'Incorrect passcode.';
        adminLoginError.classList.remove('hidden');
      }
    }
  };

  if (adminLoginButton) {
    adminLoginButton.addEventListener('click', handleAdminLogin);
  }

  if (adminPasscodeInput) {
    adminPasscodeInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        handleAdminLogin();
      }
    });
  }

  // --- Reword with AI Logic ---
  const rewordModal = document.getElementById('rewordModal');
  const rewordOptionsContainer = document.getElementById('rewordOptionsContainer');
  const confirmRewordButton = document.getElementById('confirmReword');
  const cancelRewordButton = document.getElementById('cancelReword');

  const showRewordModal = (versions: string[]) => {
    if (!rewordModal || !rewordOptionsContainer) return;
    rewordOptionsContainer.innerHTML = '';
    versions.forEach((version, index) => {
        const id = `reword-option-${index}`;
        const optionDiv = document.createElement('div');
        optionDiv.className = 'reword-option';
        
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.id = id;
        radio.name = 'rewordSelection';
        radio.value = version;
        if (index === 0) radio.checked = true;

        const label = document.createElement('label');
        label.htmlFor = id;
        label.textContent = version;

        optionDiv.appendChild(radio);
        optionDiv.appendChild(label);
        rewordOptionsContainer.appendChild(optionDiv);
    });
    rewordModal.classList.remove('hidden');
  };

  const hideRewordModal = () => {
    if(rewordModal) rewordModal.classList.add('hidden');
    targetTextareaIdForReword = null;
  };

  const handleRewordClick = async (buttonId: string, textareaId: string, context: string) => {
    const button = document.getElementById(buttonId) as HTMLButtonElement;
    const textarea = document.getElementById(textareaId) as HTMLTextAreaElement;
    const apiKey = (document.getElementById('apiKey') as HTMLInputElement)?.value;

    if (!apiKey) {
      alert('Please enter your Gemini API Key in the Admin Login tab.');
      // Switch to admin tab
      document.querySelector('.tab-button[data-tab="admin"]')?.dispatchEvent(new Event('click'));
      return;
    }
    if (!textarea.value.trim()) {
      alert('Please enter some text to reword.');
      return;
    }

    button.disabled = true;
    button.textContent = '...';

    try {
      const ai = new GoogleGenAI({ apiKey });
      const systemInstruction = "You are an expert in buried steam, hot water, and superheated hot water systems, and also an expert in natural gas buried distribution and transmission systems. You understand the critical safety issues that can arise from excessive heat transfer from a heat source line to a nearby gas line. Your task is to rephrase the user's text to be more professional, clear, and technically accurate for an engineering assessment report. Provide exactly 3 distinct, reworded versions.";
      const contents = `Please reword the following text about "${context}":\n\n"${textarea.value}"`;

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents,
        config: {
          systemInstruction,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              versions: {
                type: Type.ARRAY,
                items: { type: Type.STRING },
                description: "An array of three distinct reworded versions of the original text."
              }
            },
            required: ["versions"]
          }
        }
      });

      const jsonString = response.text.trim();
      const result = JSON.parse(jsonString);

      if (result.versions && Array.isArray(result.versions) && result.versions.length > 0) {
        targetTextareaIdForReword = textareaId;
        showRewordModal(result.versions);
      } else {
        throw new Error("Invalid response format from API.");
      }
    } catch (error) {
      console.error("AI Reword Error:", error);
      alert(`Failed to generate suggestions. Please check your API key and the browser console for details.\n${error instanceof Error ? error.message : ''}`);
    } finally {
      button.disabled = false;
      button.textContent = 'Reword';
    }
  };

  const setupRewordButtons = () => {
    const rewordTargets = [
      { btnId: 'rewordProjectDescription', txtId: 'projectDescription', context: 'the project description' },
      { btnId: 'rewordSystemDutyCycle', txtId: 'systemDutyCycle', context: "the system's duty cycle or uptime" },
      { btnId: 'rewordPipeCasingInfo', txtId: 'pipeCasingInfo', context: 'pipe casing or conduit information' },
      { btnId: 'rewordHeatLossEvidence', txtId: 'heatLossEvidence', context: 'evidence of heat loss at the surface' },
      { btnId: 'rewordInsulationCondition', txtId: 'insulationCondition', context: 'the known condition of the insulation' },
      { btnId: 'rewordAdditionalInfo', txtId: 'additionalInfo', context: 'additional information provided by the operator' },
    ];
    rewordTargets.forEach(({ btnId, txtId, context }) => {
      document.getElementById(btnId)?.addEventListener('click', () => handleRewordClick(btnId, txtId, context));
    });
  };

  const setupQuestionnaireReword = () => {
    const rewordAnswerBtn = document.getElementById('rewordAnswer');
    rewordAnswerBtn?.addEventListener('click', () => {
      const questionText = document.getElementById('question-text')?.textContent || '';
      handleRewordClick('rewordAnswer', 'question-answer', `an answer to the question: "${questionText}"`);
    });
  };

  confirmRewordButton?.addEventListener('click', () => {
    if (targetTextareaIdForReword) {
      const selectedRadio = document.querySelector<HTMLInputElement>('input[name="rewordSelection"]:checked');
      if (selectedRadio) {
        const textarea = document.getElementById(targetTextareaIdForReword) as HTMLTextAreaElement;
        textarea.value = selectedRadio.value;
        // Trigger input event to make sure auto-resize recalculates
        textarea.dispatchEvent(new Event('input', { bubbles: true }));
      }
    }
    hideRewordModal();
  });

  cancelRewordButton?.addEventListener('click', hideRewordModal);
  rewordModal?.addEventListener('click', (e) => {
    if (e.target === rewordModal) {
      hideRewordModal();
    }
  });

  // --- Image Uploader and Resizer ---
  const setupImageUploader = () => {
      const fileInput = document.getElementById('fieldPhotos') as HTMLInputElement;
      const previewContainer = document.getElementById('image-preview-container');

      if (!fileInput || !previewContainer) return;

      fileInput.addEventListener('change', () => {
          previewContainer.innerHTML = ''; // Clear previous previews
          const files = fileInput.files;
          if (!files) return;

          Array.from(files).forEach(file => {
              // The input 'accept' attribute handles the filtering, but we can double-check
              if (file.type === 'image/jpeg') {
                  const reader = new FileReader();
                  reader.onload = (e) => {
                      if (e.target?.result) {
                          createResizableImage(e.target.result as string, previewContainer);
                      }
                  };
                  reader.readAsDataURL(file);
              }
          });
      });
  };

  // --- Dynamic Recommendations ---
  const addRecommendation = (text: string = '') => {
    const container = document.getElementById('recommendationsContainer');
    if (!container) return;

    const itemId = `recommendation-${recommendationCounter}`;
    const rewordBtnId = `reword-recommendation-${recommendationCounter}`;
    
    const item = document.createElement('div');
    item.className = 'recommendation-item';

    const header = document.createElement('div');
    header.className = 'recommendation-header';

    const label = document.createElement('label');
    label.htmlFor = itemId;
    label.textContent = `Recommendation #${container.children.length + 1}`;

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'delete-recommendation-btn';
    deleteBtn.innerHTML = '&times;';
    deleteBtn.setAttribute('aria-label', 'Delete recommendation');
    deleteBtn.addEventListener('click', () => {
        item.remove();
        // Update labels of remaining items
        const remainingItems = container.querySelectorAll('.recommendation-item');
        remainingItems.forEach((remItem, index) => {
            const remLabel = remItem.querySelector('label');
            if (remLabel) {
                remLabel.textContent = `Recommendation #${index + 1}`;
            }
        });
    });

    header.appendChild(label);
    header.appendChild(deleteBtn);

    const textareaWrapper = document.createElement('div');
    textareaWrapper.className = 'textarea-wrapper';

    const textarea = document.createElement('textarea');
    textarea.id = itemId;
    textarea.className = 'recommendation-textarea';
    textarea.rows = 4;
    textarea.placeholder = 'Enter recommendation...';
    textarea.value = text;

    const adjustTextareaHeight = () => {
        textarea.style.height = 'auto';
        textarea.style.height = `${textarea.scrollHeight}px`;
    };
    textarea.addEventListener('input', adjustTextareaHeight);
    
    const rewordBtn = document.createElement('button');
    rewordBtn.type = 'button';
    rewordBtn.id = rewordBtnId;
    rewordBtn.className = 'reword-button';
    rewordBtn.textContent = 'Reword';
    rewordBtn.setAttribute('aria-label', `Reword Recommendation #${container.children.length + 1}`);
    rewordBtn.classList.toggle('hidden', !isAdminAuthenticated);
    rewordBtn.addEventListener('click', () => handleRewordClick(rewordBtn.id, textarea.id, 'an engineering recommendation'));

    textareaWrapper.appendChild(textarea);
    textareaWrapper.appendChild(rewordBtn);

    item.appendChild(header);
    item.appendChild(textareaWrapper);

    container.appendChild(item);
    
    // Initial resize
    setTimeout(adjustTextareaHeight, 0);

    recommendationCounter++;
  };

  // --- Initialization ---
  updateEvaluatorInputs(); // Initial call for evaluator inputs
  toggleConditionalFields(); // Initial call to set up the form state based on default selections
  handleGasOperatorNameChange(); // Set initial state for gas operator name field
  handleGasPipeMaterialChange(); // Set initial state for gas material dependent fields
  handleGasSdrChange(); // Set initial state for custom SDR input
  handleGasLineOrientationChange(); // Set initial state for gas orientation fields
  handleHeatLossSourceChange(); // Set initial state for heat loss evidence fields
  handleInsulationConditionSourceChange(); // Set initial state for insulation condition fields
  setupRewordButtons();
  setupQuestionnaireReword();
  setupImageUploader();
  addRecommendation(); // Add one initial recommendation box
  document.getElementById('addRecommendationButton')?.addEventListener('click', () => addRecommendation());

  // --- Save/Load Functionality ---
  const saveButton = document.getElementById('saveButton');
  const loadButton = document.getElementById('loadButton');
  const loadFileInput = document.getElementById('loadFile') as HTMLInputElement;

  const handleSave = () => {
    const formData: { [key: string]: any } = {};
    const form = document.querySelector('.project-info-form');
    if (!form) return;

    const elements = form.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    
    elements.forEach(el => {
      const id = el.id;
      const name = el.name;
      const type = (el as any).type;

      if (type === 'radio') {
        if ((el as HTMLInputElement).checked) {
          formData[name] = (el as HTMLInputElement).value;
        }
      } else if (el.tagName.toLowerCase() === 'select' && (el as HTMLSelectElement).multiple) {
        if (id) {
            formData[id] = Array.from((el as HTMLSelectElement).selectedOptions).map(opt => opt.value);
        }
      } else if (id && type !== 'file') {
        formData[id] = el.value;
      }
    });
    
    const fullData = getFormData(); // Get data including dynamic fields

    const reportHTML = document.getElementById('generated-report-container')?.innerHTML || '';
    const saveData = {
        formData: fullData, // Save the full data object
        reportHTML,
        questionAnswers
    };

    const dataStr = JSON.stringify(saveData, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `assessment-${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleFileSelect = (event: Event) => {
    const input = event.target as HTMLInputElement;
    if (!input.files || input.files.length === 0) return;

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result as string);
        populateForm(data);
      } catch (error) {
        console.error('Error parsing JSON file:', error);
        alert('Could not load assessment. The file is not a valid JSON file.');
      }
    };

    reader.onerror = () => {
      console.error('Error reading file:', reader.error);
      alert('An error occurred while reading the file.');
    };

    reader.readAsText(file);
    input.value = ''; // Reset file input
  };

  const populateForm = (savedData: { [key: string]: any }) => {
    const data = savedData.formData;
    if (!data) {
        alert('Invalid save file format.');
        return;
    }

    for (const key in data) {
      if (typeof data[key] === 'string' || typeof data[key] === 'number') {
        const el = document.getElementById(key) as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement;
        if (el) {
            el.value = String(data[key]);
        }
      }
    }
    
    // Special handling for evaluator names
    if (data.evaluatorNames && Array.isArray(data.evaluatorNames)) {
        const numEvaluatorsSelect = document.getElementById('numEvaluators') as HTMLSelectElement;
        if (numEvaluatorsSelect) {
            const num = data.evaluatorNames.length;
            if (num > 0 && num <= 3) {
                numEvaluatorsSelect.value = String(num);
                // This will trigger the input creation
                numEvaluatorsSelect.dispatchEvent(new Event('change', { bubbles: true }));

                // Use setTimeout to allow the DOM to update with the new inputs
                setTimeout(() => {
                    const nameInputs = document.querySelectorAll<HTMLInputElement>('.evaluator-name-input');
                    data.evaluatorNames.forEach((name: string, index: number) => {
                        if (nameInputs[index]) {
                            nameInputs[index].value = name;
                        }
                    });
                }, 0);
            }
        }
    }
    
    // Special handling for heat source type radio and custom input
    if (data.heatSourceType) {
      const heatSourceRadios = document.querySelectorAll<HTMLInputElement>('input[name="heatSourceType"]');
      let foundMatch = false;

      // Check for 'Other' case first
      if (data.heatSourceType.startsWith('Other:')) {
        const otherRadio = document.getElementById('other') as HTMLInputElement;
        if (otherRadio) {
          otherRadio.checked = true;
          foundMatch = true;
          const customInput = document.getElementById('customHeatSourceType') as HTMLInputElement;
          const customValue = data.heatSourceType.substring('Other:'.length).trim();
          if (customInput && customValue !== 'not specified') {
            customInput.value = customValue;
          }
        }
      }

      // If not 'Other', try to match other labels
      if (!foundMatch) {
        heatSourceRadios.forEach(radio => {
          const label = document.querySelector(`label[for="${radio.id}"]`);
          if (label && label.textContent === data.heatSourceType) {
            radio.checked = true;
          }
        });
      }
    }

    // Handle Radios by checking against both value and label text to support inconsistent save data
    const radioNames = new Set(Array.from(document.querySelectorAll<HTMLInputElement>('input[type="radio"]')).map(r => r.name));
    radioNames.forEach(name => {
      if (data.hasOwnProperty(name) && name !== 'heatSourceType') { // heatSourceType is handled specially
        const savedValue = data[name];
        const radios = document.querySelectorAll<HTMLInputElement>(`input[name="${name}"]`);
        
        // Handle "Other: " combined strings first
        if (typeof savedValue === 'string' && savedValue.startsWith('Other:')) {
            const otherRadio = document.querySelector<HTMLInputElement>(`input[name="${name}"][value="other"]`);
            if (otherRadio) {
                otherRadio.checked = true;
                const customInputId = {
                    'heatLossEvidenceSource': 'customHeatLossSource',
                    'insulationConditionSource': 'customInsulationConditionSource'
                }[name];
                if (customInputId) {
                    const customInput = document.getElementById(customInputId) as HTMLInputElement;
                    if (customInput) {
                        customInput.value = savedValue.substring('Other:'.length).trim();
                    }
                }
            }
        } else {
            radios.forEach(radio => {
                const label = document.querySelector(`label[for="${radio.id}"]`);
                // Check against value OR label text content
                if (radio.value === savedValue || (label && label.textContent === savedValue)) {
                    radio.checked = true;
                }
            });
        }
      }
    });


    // Handle multiselect for connection types
    if (data.connectionsValue && Array.isArray(data.connectionsValue)) {
      const selectEl = document.getElementById('connectionTypes') as HTMLSelectElement;
      if (selectEl) {
        let isOtherSelected = false;
        let customText = '';
        const otherEntry = data.connectionsValue.find((v: string) => v.startsWith('Other:'));
        if (otherEntry) {
            isOtherSelected = true;
            const text = otherEntry.substring('Other:'.length).trim();
            if (text !== 'not specified') {
                customText = text;
            }
        }

        const valuesToSelect = data.connectionsValue.filter((v: string) => !v.startsWith('Other:'));

        Array.from(selectEl.options).forEach(option => {
            if (option.value === 'other') {
                option.selected = isOtherSelected;
            } else {
                option.selected = valuesToSelect.includes(option.text);
            }
        });
        
        const customInput = document.getElementById('customConnectionTypes') as HTMLTextAreaElement;
        if (customInput) {
            customInput.value = customText;
        }
      }
    }
    
    // Special handling for wall thickness to re-populate options first
    const pipelineDiameterSelect = document.getElementById('pipelineDiameter') as HTMLSelectElement;
    if(pipelineDiameterSelect) {
      pipelineDiameterSelect.dispatchEvent(new Event('change'));
      setTimeout(() => {
        const wallThicknessSelect = document.getElementById('wallThickness') as HTMLSelectElement;
        if (wallThicknessSelect && data['wallThickness']) {
          wallThicknessSelect.value = data['wallThickness'];
          wallThicknessSelect.dispatchEvent(new Event('change'));
        }
        const customWallThickness = document.getElementById('customWallThickness') as HTMLInputElement;
        if (customWallThickness && data['customWallThickness']) {
            customWallThickness.value = data['customWallThickness'];
        }
      }, 0);
    }
    
    // Special handling for parallel coordinates
    if (data.gasLineOrientation === 'Parallel' || data.parallelLength !== 'N/A') {
        const parallelLengthInput = document.getElementById('parallelLength') as HTMLInputElement;
        if (parallelLengthInput && data.parallelLength) {
            parallelLengthInput.value = data.parallelLength;
            parallelLengthInput.dispatchEvent(new Event('input', { bubbles: true }));

            setTimeout(() => {
                if (data.parallelCoordinates && Array.isArray(data.parallelCoordinates)) {
                    const savedCoords = data.parallelCoordinates;
                    const latInputs = document.querySelectorAll<HTMLInputElement>('#parallelCoordinatesWrapper input[id^="lat-parallel-"]');
                    
                    latInputs.forEach((latInput, i) => {
                        const index = latInput.id.split('-').pop();
                        const lngInput = document.getElementById(`lng-parallel-${index}`) as HTMLInputElement;

                        if (savedCoords[i] && lngInput) {
                            latInput.value = savedCoords[i].lat || '';
                            lngInput.value = savedCoords[i].lng || '';
                        }
                    });
                }
            }, 100); // Increased timeout for DOM update
        }
    }

    // Clear existing dynamic recommendations before populating
    const recContainer = document.getElementById('recommendationsContainer');
    if (recContainer) recContainer.innerHTML = '';

    // Repopulate recommendations
    if (data.recommendations && Array.isArray(data.recommendations) && data.recommendations.length > 0) {
        data.recommendations.forEach((recText: string) => {
            addRecommendation(recText);
        });
    } else if (data.recommendationsText) { // Handle legacy save files
        addRecommendation(data.recommendationsText);
    }
    else {
        // For new forms or empty saves, add one empty box
        addRecommendation('');
    }


    // Trigger change events to update UI
    const elementsToUpdate = new Set<HTMLElement | null>([
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="heatSourceType"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="gasLineOrientation"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="heatLossEvidenceSource"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="insulationConditionSource"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="heatSourceBeddingType"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="heatSourceBeddingUseCustomK"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="gasLineBeddingType"]')),
      ...Array.from(document.querySelectorAll<HTMLInputElement>('input[name="gasBeddingUseCustomK"]')),
      document.getElementById('pipelineDiameter'),
      document.getElementById('gasPipelineDiameter'),
      document.getElementById('pipeMaterial'),
      document.getElementById('gasPipeMaterial'),
      document.getElementById('wallThickness'), // Needs another trigger after population
      document.getElementById('pipeInsulationType'),
      document.getElementById('connectionTypes'),
      document.getElementById('gasOperatorName'),
      document.getElementById('gasCoatingType'),
      document.getElementById('soilType')
    ]);
    
    elementsToUpdate.forEach(el => {
        if (el) {
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }
    });

    // Handle loading gas wall thickness and SDR after a delay
    setTimeout(() => {
        if (data.gasPipeWallThickness && data.gasPipeWallThickness !== 'N/A') {
            const gasWallThicknessSelect = document.getElementById('gasWallThickness') as HTMLSelectElement;
            const customGasWallThicknessInput = document.getElementById('customGasWallThickness') as HTMLInputElement;
            
            // Check if the value exists in the standard options
            const optionExists = Array.from(gasWallThicknessSelect.options).some(opt => opt.value === data.gasPipeWallThickness);

            if (optionExists) {
                gasWallThicknessSelect.value = data.gasPipeWallThickness;
            } else {
                // If it doesn't exist, it's a custom value
                gasWallThicknessSelect.value = 'custom';
                customGasWallThicknessInput.value = data.gasPipeWallThickness;
            }
            gasWallThicknessSelect.dispatchEvent(new Event('change', { bubbles: true }));
        }

        if (data.gasPipeSDR && data.gasPipeSDR !== 'N/A') {
            const gasSdrSelect = document.getElementById('gasSdr') as HTMLSelectElement;
            const customGasSdrInput = document.getElementById('customGasSdr') as HTMLInputElement;

            const optionExists = Array.from(gasSdrSelect.options).some(opt => opt.value === data.gasPipeSDR);

            if (optionExists) {
                gasSdrSelect.value = data.gasPipeSDR;
            } else {
                gasSdrSelect.value = 'other';
                customGasSdrInput.value = data.gasPipeSDR;
            }
            gasSdrSelect.dispatchEvent(new Event('change', { bubbles: true }));
        }
    }, 100);

    // Repopulate final report
    if (savedData.reportHTML) {
      const reportContainer = document.getElementById('generated-report-container');
      const reportActions = document.getElementById('report-actions');
      if (reportContainer) {
        reportContainer.innerHTML = savedData.reportHTML;
        if(savedData.reportHTML.trim()) {
            reportActions?.classList.remove('hidden');
        }
      }
    }

    if (savedData.questionAnswers) {
        questionAnswers = savedData.questionAnswers;
    }
  };

  if(saveButton) saveButton.addEventListener('click', handleSave);
  if(loadButton) loadButton.addEventListener('click', () => loadFileInput.click());
  if(loadFileInput) loadFileInput.addEventListener('change', handleFileSelect);

  // --- Final Report Questionnaire Logic ---
  const questionnaireWrapper = document.getElementById('report-questionnaire');
  const generationPromptWrapper = document.getElementById('report-generation-prompt');
  const reportAdminMessage = document.getElementById('report-admin-only-message');
  const questionTextEl = document.getElementById('question-text');
  const questionAnswerTextarea = document.getElementById('question-answer') as HTMLTextAreaElement;
  const nextQuestionBtn = document.getElementById('next-question-btn') as HTMLButtonElement;
  const skipQuestionBtn = document.getElementById('skip-question-btn');
  const clarifyQuestionBtn = document.getElementById('clarify-question-btn') as HTMLButtonElement | null;
  const currentQNumEl = document.getElementById('current-q-num');
  const totalQNumEl = document.getElementById('total-q-num');
  const progressBar = document.getElementById('progress-bar');
  const generateReportBtn = document.getElementById('generate-report-btn');
  
  const handleFinalReportTabClick = () => {
    if (!isAdminAuthenticated) {
      reportAdminMessage?.classList.remove('hidden');
      questionnaireWrapper?.classList.add('hidden');
      generationPromptWrapper?.classList.add('hidden');
      return;
    }

    reportAdminMessage?.classList.add('hidden');

    if (!lastCalculationResults) {
      updateCalculationSummary();
      const confirmation = confirm("You must run a calculation before generating a final report. Would you like to switch to the Calculation tab now?");
      if (confirmation) {
        document.querySelector('.tab-button[data-tab="calculation"]')?.dispatchEvent(new Event('click'));
      }
      return;
    }
    
    // Check if questionnaire is done
    if (currentQuestionIndex >= reportQuestions.length) {
      questionnaireWrapper?.classList.add('hidden');
      generationPromptWrapper?.classList.remove('hidden');
    } else {
      questionnaireWrapper?.classList.remove('hidden');
      generationPromptWrapper?.classList.add('hidden');
      // If starting for the first time
      if (currentQuestionIndex === -1) {
        startQuestionnaire();
      }
    }
  };

  const startQuestionnaire = () => {
    currentQuestionIndex = 0;
    questionAnswers = [];
    displayCurrentQuestion();
  };

  const displayCurrentQuestion = () => {
    if (currentQuestionIndex >= reportQuestions.length) {
      questionnaireWrapper?.classList.add('hidden');
      generationPromptWrapper?.classList.remove('hidden');
      return;
    }
    
    if(questionTextEl) questionTextEl.textContent = reportQuestions[currentQuestionIndex];
    questionAnswerTextarea.value = ''; // Clear previous answer
    questionAnswerTextarea.dispatchEvent(new Event('input')); // for auto-resize
    if(currentQNumEl) currentQNumEl.textContent = String(currentQuestionIndex + 1);
    if(totalQNumEl) totalQNumEl.textContent = String(reportQuestions.length);

    if (progressBar) {
      const progress = ((currentQuestionIndex) / reportQuestions.length) * 100;
      progressBar.style.width = `${progress}%`;
    }

    if(nextQuestionBtn) {
        if (currentQuestionIndex === reportQuestions.length - 1) {
            nextQuestionBtn.textContent = 'Finish';
        } else {
            nextQuestionBtn.textContent = 'Next';
        }
    }
  };

  const saveAndGoToNext = () => {
    const answer = questionAnswerTextarea.value.trim();
    if (answer) {
      questionAnswers.push({
        question: reportQuestions[currentQuestionIndex],
        answer: answer,
        originalQuestion: reportQuestions[currentQuestionIndex]
      });
    }
    currentQuestionIndex++;
    displayCurrentQuestion();
  };

  const handleSkip = () => {
    currentQuestionIndex++;
    displayCurrentQuestion();
  };

  const handleClarify = async () => {
      const apiKey = (document.getElementById('apiKey') as HTMLInputElement)?.value;
      if (!apiKey) {
          alert('Please enter your Gemini API Key in the Admin Login tab to use this feature.');
          document.querySelector('.tab-button[data-tab="admin"]')?.dispatchEvent(new Event('click'));
          return;
      }
      if (!clarifyQuestionBtn) return;
      
      clarifyQuestionBtn.textContent = '...';
      (clarifyQuestionBtn as HTMLButtonElement).disabled = true;

      try {
          const ai = new GoogleGenAI({ apiKey });
          const currentQuestion = reportQuestions[currentQuestionIndex];
          const systemInstruction = "You are an expert in civil and mechanical engineering, specializing in underground utilities like gas and steam lines. Your task is to clarify an engineering question by rephrasing it in a simpler, more direct way, and by providing brief examples of the type of information being sought. Do this in one single paragraph.";
          const contents = `Clarify this question for an engineer filling out an assessment form: "${currentQuestion}"`;
          
          const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents,
            config: { systemInstruction }
          });
          
          const clarification = response.text !== undefined ? response.text : 'Could not get clarification.';
          alert(`Clarification:\n\n${clarification}`);
      } catch (error) {
          console.error("AI Clarification Error:", error);
          alert(`Failed to get clarification. Please check your API key and the browser console.\n${error instanceof Error ? error.message : ''}`);
      } finally {
          clarifyQuestionBtn.textContent = 'Clarify';
          (clarifyQuestionBtn as HTMLButtonElement).disabled = false;
      }
  };

  nextQuestionBtn?.addEventListener('click', saveAndGoToNext);
  skipQuestionBtn?.addEventListener('click', handleSkip);
  clarifyQuestionBtn?.addEventListener('click', handleClarify);

  // --- Final Report Generation ---
  const generateFinalReport = async () => {
    const apiKey = (document.getElementById('apiKey') as HTMLInputElement)?.value;
    if (!apiKey) {
      alert('Please enter your Gemini API Key in the Admin Login tab to generate the report.');
      document.querySelector('.tab-button[data-tab="admin"]')?.dispatchEvent(new Event('click'));
      return;
    }

    const spinner = document.getElementById('report-loading-spinner');
    const reportContainer = document.getElementById('generated-report-container');
    const reportActions = document.getElementById('report-actions');

    if(spinner) spinner.classList.remove('hidden');
    if(reportContainer) reportContainer.innerHTML = '';
    if(reportActions) reportActions.classList.add('hidden');
    
    try {
      const ai = new GoogleGenAI({ apiKey });
      const formData = getFormData();
      
      let systemInstruction = `You are a professional engineering consultant specializing in the safety and integrity of underground utility infrastructure, particularly natural gas distribution systems and high-temperature pipelines (steam, hot water). Your task is to generate a formal, comprehensive "Thermal–Gas Line Encroachment Assessment Report" based on the provided data.

      The report must be structured with the following sections, in this exact order:
      1.  **Executive Summary:** Start with a brief, high-level overview. State the purpose of the assessment, the location, and the single most critical finding (e.g., "the calculated gas line temperature is [X]°F, which is [above/below] the material's maximum allowable limit of [Y]°F"). Conclude with a summary of the primary recommendation.
      2.  **Introduction:** Detail the project name, location, date of assessment, and the names of the evaluators. Provide a clear description of the project's scope and objectives.
      3.  **Data Summary:** Present all the input data in a structured, easy-to-read format. Use tables for clarity where appropriate. Create distinct subsections for "Heat Source Data," "Gas Line Data," "Soil & Bedding Data," and "Field Visit Observations." You MUST include all provided data points; if a value is "N/A" or "Unknown," state that explicitly. This includes all operator contact information, 811 status, confirmation dates, etc. The "Field Visit" section is critical and must include all notes from the on-site visit.
      4.  **Analysis & Calculation Results:** This is a critical section. First, explain the methodology used for the thermal analysis in simple terms (e.g., "A steady-state heat transfer model based on thermal resistance was used..."). Then, present the key calculated results clearly, especially the "Calculated Temperature at Gas Line Surface." Compare this calculated temperature directly to the gas line's maximum allowable temperature.
          - If the gas line is **coated steel**, use the specific coating's maximum temperature limit for this comparison.
          - If the gas line is **plastic (HDPE, MDPE, etc.)**, you MUST state that while the material has a higher nominal temperature limit (e.g., 140°F), a conservative **operational limit of 70°F** is being enforced for this assessment to ensure long-term safety and prevent material degradation. The primary comparison, risk assessment, and conclusions MUST be based on this 70°F limit. Mention the material's actual limits for context but emphasize the 70°F operational constraint.
      5.  **Risk Assessment:** Based on the analysis, evaluate the potential risks. Synthesize information from the questionnaire answers, field visit observations, and the calculated results. Discuss the likelihood and consequences of the gas line's temperature exceeding its limit, referencing relevant standards like 49 CFR 192, ASME B31.8, and operator-specific guidelines. Consider factors like material degradation (e.g., PE softening, coating disbondment), potential for pressure de-rating, and public safety implications.
      6.  **Conclusions:** Summarize the findings of the assessment in a clear, numbered list. Each conclusion must be a direct result of the data and analysis.
      7.  **Recommendations:** This is the most important section. Provide a numbered list of clear, actionable recommendations. Each recommendation must be justified by the findings and directly address the identified risks. **You must professionally rephrase and integrate all of the user-provided recommendations from the form data's "recommendations" array.** These are non-negotiable and must be included. Then, generate any additional recommendations you deem necessary based on the full data set. Reference relevant regulations (49 CFR 192, ASME B31.8, GTPC Guidelines) to support your recommendations.
      
      **Formatting and Tone:**
      - Use a formal, objective, and professional tone suitable for an engineering report.
      - Use Markdown for formatting (headings, bold text, bullet points, tables).
      - Ensure all numerical values are presented with their correct units (e.g., °F, psig, feet, inches).
      - Be precise and unambiguous in your language.
      - Convert the final markdown to HTML.
      - Use a clean, professional HTML structure with appropriate tags (<h1>, <h2>, <h3>, <p>, <ul>, <li>, <table>, <th>, <tr>, <td>).`;
      
      let contents = `
        Here is the data for the assessment:
        - **Form Data:** ${JSON.stringify(formData, null, 2)}
        - **Calculation Results:** ${JSON.stringify(lastCalculationResults, null, 2)}
        - **Questionnaire Answers:** ${JSON.stringify(questionAnswers, null, 2)}
        
        Please generate the report.
      `;

      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents,
        config: { systemInstruction }
      });
      
      const reportHtml = response.text ?? '';

      if(reportContainer) reportContainer.innerHTML = reportHtml;
      if(reportActions) reportActions.classList.remove('hidden');

    } catch(error) {
        console.error("Final Report Generation Error:", error);
        if (reportContainer) {
          reportContainer.innerHTML = `<div class="calculation-error"><h3>Report Generation Failed</h3><p>An error occurred while generating the report. Please check your API key and the browser console for details.</p><p>${error instanceof Error ? error.message : 'Unknown error'}</p></div>`;
        }
    } finally {
        if(spinner) spinner.classList.add('hidden');
    }
  };

  generateReportBtn?.addEventListener('click', generateFinalReport);

  document.getElementById('send-to-pdf-btn')?.addEventListener('click', () => {
    // Temporarily remove the print button from the DOM before printing
    const reportActions = document.getElementById('report-actions');
    const wasHidden = reportActions?.classList.contains('hidden');
    reportActions?.classList.add('hidden');

    window.print();
    
    // Add the button back after printing dialog is closed
    if(reportActions && !wasHidden) {
        reportActions.classList.remove('hidden');
    }
  });


  // --- Calculation Logic ---
  document.getElementById('calculateButton')?.addEventListener('click', () => {
    const data = getFormData();
    const resultsContainer = document.getElementById('resultsContainer')!;
    resultsContainer.innerHTML = '';
    
    const errors: string[] = [];
    
    // Validation
    if (!data.isHeatSourceApplicable) {
      errors.push("No applicable heat source selected in the 'Heat Source Data' tab. Cannot perform calculation.");
    }
    if (!data.maxTemp) errors.push("Heat Source Max Operating Temperature is required.");
    if (!data.pipelineDiameter || data.pipelineDiameter === 'N/A') errors.push("Heat Source Pipeline Nominal Diameter is required.");
    if (!data.heatSourcePipeMaterial || data.heatSourcePipeMaterial === 'N/A') errors.push("Heat Source Pipe Material is required.");
    if (!data.heatSourceWallThickness || data.heatSourceWallThickness === 'N/A') errors.push("Heat Source Pipe Wall Thickness is required.");
    if (data.insulationType !== 'None' && (!data.insulationThickness)) {
      errors.push("Insulation Thickness is required when insulation is specified.");
    }
    if (!data.heatSourceDepth) errors.push("Heat Source Depth is required.");
    if (!data.gasPipelineDiameter || data.gasPipelineDiameter === 'N/A') errors.push("Gas Line Pipeline Nominal Diameter is required.");
    if (!data.gasPipeMaterial || data.gasPipeMaterial === 'N/A') errors.push("Gas Pipeline Material is required.");
    if (!data.depthOfBurialGasLine) errors.push("Depth of Gas Line is required.");
    if (data.gasLineOrientation === 'Parallel' && !data.parallelDistance) {
      errors.push("Distance between lines is required for parallel orientation.");
    }
    if (!data.soilType || data.soilType === 'N/A') errors.push("Native Soil Type is required.");
    if (!data.soilThermalConductivity) errors.push("Native Soil Thermal Conductivity is required.");
    if (!data.groundSurfaceTemperature) errors.push("Average Ground Surface Temperature is required.");

    if (errors.length > 0) {
      const errorList = errors.map(e => `<li>${e}</li>`).join('');
      resultsContainer.innerHTML = `<div class="calculation-error"><h3>Missing Information</h3><p>Please provide the following required information to perform the calculation:</p><ul>${errorList}</ul></div>`;
      resultsContainer.classList.remove('hidden');
      return;
    }

    try {
      // --- Calculations ---
      const T_hs = parseFloat(data.maxTemp);
      const T_surface = parseFloat(data.groundSurfaceTemperature);
      
      const D_hs_nominal_in = parseFloat(data.pipelineDiameter);
      const D_hs_outer_in = npsToOdMapping[String(D_hs_nominal_in)];
      const t_hs_wall_in = parseFloat(data.heatSourceWallThickness.replace(/[^0-9.]/g, ''));
      const D_hs_inner_in = D_hs_outer_in - (2 * t_hs_wall_in);
      
      const D_hs_outer_ft = D_hs_outer_in / 12;
      const D_hs_inner_ft = D_hs_inner_in / 12;

      const k_pipe_material_str = data.heatSourcePipeMaterial;
      let k_pipe: number;
      if(k_pipe_material_str.startsWith('Other:')) {
          k_pipe = parseFloat((document.getElementById('customThermalConductivity') as HTMLInputElement).value);
          if(isNaN(k_pipe)) throw new Error("Custom pipe material thermal conductivity is not a valid number.");
      } else {
          const materialKey = (document.getElementById('pipeMaterial') as HTMLSelectElement).value;
          k_pipe = pipeMaterialData[materialKey as keyof typeof pipeMaterialData].thermalConductivity;
      }
      
      const R_pipe_wall = Math.log(D_hs_outer_ft / D_hs_inner_ft) / (2 * Math.PI * k_pipe);
      
      let R_insulation = 0;
      let D_ins_outer_ft: number | undefined;

      if (data.insulationType !== 'None') {
        const t_ins_in = parseFloat(data.insulationThickness);
        D_ins_outer_ft = (D_hs_outer_in + (2 * t_ins_in)) / 12;
        
        let k_ins: number;
        if(data.insulationType.startsWith('Other:')) {
            k_ins = parseFloat(data.customInsulationThermalConductivity);
            if(isNaN(k_ins)) throw new Error("Custom insulation thermal conductivity is not a valid number.");
        } else {
            const heatSourceType = document.querySelector<HTMLInputElement>('input[name="heatSourceType"]:checked')?.value;
            const insulationKey = (document.getElementById('pipeInsulationType') as HTMLSelectElement).value;
            if(heatSourceType === 'steam' || heatSourceType === 'superHeatedHotWater') {
              k_ins = insulationData.steam[insulationKey].thermalConductivity;
            } else if (heatSourceType === 'hotWater') {
              k_ins = insulationData.hotWater[insulationKey].thermalConductivity;
            } else {
                throw new Error("Could not determine insulation thermal conductivity for the selected heat source.");
            }
        }
        R_insulation = Math.log(D_ins_outer_ft / D_hs_outer_ft) / (2 * Math.PI * k_ins);
      }
      
      const Z_hs_ft = parseFloat(data.heatSourceDepth);
      const k_soil = parseFloat(data.soilThermalConductivity);
      const R_surface_ft = data.insulationType !== 'None' ? D_ins_outer_ft! / 2 : D_hs_outer_ft / 2;
      
      const R_soil_hs = Math.log((2 * Z_hs_ft - R_surface_ft) / R_surface_ft) / (2 * Math.PI * k_soil);
      
      const R_total = R_pipe_wall + R_insulation + R_soil_hs;
      
      const delta_T = T_hs - T_surface;
      const Q = delta_T / R_total; // Heat loss per unit length (BTU/hr·ft)
      
      const D_gas_nominal_in = parseFloat(data.gasPipelineDiameter);
      const D_gas_od_in = getGasPipeOd(String(D_gas_nominal_in), data.gasPipeSizingStandard);
      const D_gas_od_ft = D_gas_od_in / 12;
      const Z_gas_ft = parseFloat(data.depthOfBurialGasLine);

      // --- Calculate Temperature at Gas Line ---
      let T_gas_line: number;

      if(data.gasLineOrientation === 'Perpendicular') {
          // CORRECTED: Using the method of images for a line source at the intersection point.
          // The temperature T at point (x,y) from a line source at (0, Z_hs) is:
          // T(x,y) = T_surface + (Q / 4*pi*k) * ln( (x^2+(y+Z_hs)^2) / (x^2+(y-Z_hs)^2) )
          // At the intersection point, x=0, and the gas line is at depth y=Z_gas_ft.
          // The formula simplifies to T = T_surface + (Q / 2*pi*k) * ln( (Z_hs+Z_gas) / |Z_hs-Z_gas| )
          const term1 = Z_hs_ft + Z_gas_ft;
          const term2 = Math.abs(Z_hs_ft - Z_gas_ft);
          
          // Prevent division by zero or log of infinity if depths are identical.
          // This indicates an intersection, a critical failure state. The temperature
          // would theoretically be infinite. We'll cap it at the heat source temp as a
          // conservative, worst-case estimation for a near-contact scenario.
          if (term2 < 0.001) { // Use a small epsilon for float comparison
             T_gas_line = T_hs;
          } else {
             T_gas_line = T_surface + (Q / (2 * Math.PI * k_soil)) * Math.log(term1 / term2);
          }
      } else { // Parallel
          // The parallel calculation uses the correct method of images, accounting for the
          // distance to the image source vs. the real source. It also refines the calculation by
          // finding the temperature at the edge of the gas pipe closest to the heat source for a
          // worst-case analysis at that point.
          const C_ft = parseFloat(data.parallelDistance);
          const x_target_edge = C_ft - (D_gas_od_ft / 2);
          
          // Distance from gas pipe's closest edge to the real heat source center
          const dist_to_source = Math.sqrt(Math.pow(x_target_edge, 2) + Math.pow(Z_gas_ft - Z_hs_ft, 2));
          // Distance from gas pipe's closest edge to the image heat source center (image is at -Z_hs_ft)
          const dist_to_image = Math.sqrt(Math.pow(x_target_edge, 2) + Math.pow(Z_gas_ft + Z_hs_ft, 2));

          // Prevent division by zero if pipes are at same location.
          if (dist_to_source < 0.001) {
            T_gas_line = T_hs;
          } else {
            T_gas_line = T_surface + (Q / (2 * Math.PI * k_soil)) * Math.log(dist_to_image / dist_to_source);
          }
      }

      const tempRise = T_gas_line - T_surface;
      const heatFlux = Q / (Math.PI * D_gas_od_ft);

      // --- Store results for final report ---
      const scenarioDetails: ScenarioDetails = {
        Q,
        T_gas_line,
        delta_T,
        R_pipe_wall,
        R_insulation,
        R_soil_hs,
        R_total,
        D_hs_outer_ft,
        D_hs_inner_ft,
        R_surface_ft: R_surface_ft,
        x_target_edge: 0,
        y_target_edge: 0
      };
      if (data.insulationType !== 'None') scenarioDetails.D_ins_outer_ft = D_ins_outer_ft;
      
      if (data.gasLineOrientation === 'Parallel') {
        const C_ft = parseFloat(data.parallelDistance);
        const x_target_edge = C_ft - (D_gas_od_ft / 2);
        const y_target_edge = Z_hs_ft - Z_gas_ft;
        const r_equiv_ft = Math.sqrt(Math.pow(x_target_edge, 2) + Math.pow(y_target_edge, 2));
        scenarioDetails.x_target_edge = x_target_edge;
        scenarioDetails.y_target_edge = y_target_edge;
        scenarioDetails.r_equiv_ft = r_equiv_ft;
      }

      lastCalculationResults = {
          'Calculated Temperature at Gas Line Surface (°F)': T_gas_line.toFixed(2),
          'Temperature Rise Above Ground (°F)': tempRise.toFixed(2),
          'Heat Flux at Gas Pipe Surface (BTU/hr·ft²)': heatFlux.toFixed(2),
          'Heat Loss per Unit Length (Q, BTU/hr·ft)': Q.toFixed(2),
          'Total Thermal Resistance (R_total, hr·ft·°F/BTU)': R_total.toFixed(4),
          'Pipe Wall Thermal Resistance (hr·ft·°F/BTU)': R_pipe_wall.toFixed(4),
          'Insulation Thermal Resistance (hr·ft·°F/BTU)': R_insulation.toFixed(4),
          'Soil Thermal Resistance (hr·ft·°F/BTU)': R_soil_hs.toFixed(4),
      };

      // --- Display Results ---
      let resultsHTML = `<h3>Calculation Results</h3>`;

      // Check for temperature warning
      let tempWarningMessage = '';
      const gasMaterialValue = data.gasPipeMaterialValue;
      if (gasMaterialValue === 'coated-steel-protected' || gasMaterialValue === 'coated-steel-unprotected') {
          const maxTempStr = data.gasCoatingMaxTemp;
          if (maxTempStr && maxTempStr !== 'N/A' && maxTempStr !== 'Custom') {
              const maxAllowableTemp = parseFloat(maxTempStr);
              if (T_gas_line > maxAllowableTemp) {
                  tempWarningMessage = `<div class="calculation-warning">
                      <strong>Warning:</strong> The calculated gas line temperature of <strong>${T_gas_line.toFixed(2)}°F</strong> exceeds the maximum allowable temperature of <strong>${maxAllowableTemp}°F</strong> for the selected '${data.gasCoatingType}' coating.
                  </div>`;
              }
          }
      } else if (['hdpe', 'mdpe', 'aldyl'].includes(gasMaterialValue)) {
          const operationalLimit = 70;
          const continuousLimit = data.gasPipeContinuousLimit || '140';
          if (T_gas_line > operationalLimit) {
              tempWarningMessage = `<div class="calculation-warning">
                  <strong>Warning:</strong> The calculated gas line temperature of <strong>${T_gas_line.toFixed(2)}°F</strong> exceeds the maximum operational limit of <strong>${operationalLimit}°F</strong> for plastic pipes.
                  <p style="margin-top: 0.5rem; font-size: 0.9em; font-weight: 400;">This is a conservative operational limit to ensure long-term integrity and prevent material degradation. The nominal continuous temperature limit for this material is ${continuousLimit}°F.</p>
              </div>`;
          }
      }
      
      resultsHTML += tempWarningMessage;

      resultsHTML += `<div class="result-group">
        <div class="result-item"><span class="result-label">Calculated Temperature at Gas Line Surface:</span><span class="result-value">${T_gas_line.toFixed(2)} °F</span></div>
        <div class="result-item"><span class="result-label">Temperature Rise Above Ground:</span><span class="result-value">${tempRise.toFixed(2)} °F</span></div>
        <div class="result-item"><span class="result-label">Heat Flux at Gas Pipe Surface:</span><span class="result-value">${heatFlux.toFixed(2)} BTU/hr·ft²</span></div>
      </div>`;

      resultsHTML += `<div class="result-group">
        <h4>Intermediate Values</h4>
        <div class="result-item"><span class="result-label">Heat Loss per Unit Length (Q):</span><span>${Q.toFixed(2)} BTU/hr·ft</span></div>
        <div class="result-item"><span class="result-label">Total Thermal Resistance (R_total):</span><span>${R_total.toFixed(4)} hr·ft·°F/BTU</span></div>
        <div class="result-item"><span class="result-label">&nbsp;&nbsp;&nbsp;Pipe Wall Resistance (R_pipe_wall):</span><span>${R_pipe_wall.toFixed(4)} hr·ft·°F/BTU</span></div>
        <div class="result-item"><span class="result-label">&nbsp;&nbsp;&nbsp;Insulation Resistance (R_insulation):</span><span>${R_insulation.toFixed(4)} hr·ft·°F/BTU</span></div>
        <div class="result-item"><span class="result-label">&nbsp;&nbsp;&nbsp;Soil Resistance (R_soil_hs):</span><span>${R_soil_hs.toFixed(4)} hr·ft·°F/BTU</span></div>
      </div>`;
      
      if (data.insulationType !== 'None') {
          resultsHTML += `<div class="insulation-summary">
              <h4>Insulation Impact Summary</h4>
              <div class="summary-fact">
                  <span class="summary-fact-label">Heat Loss without Insulation (Q):</span>
                  <span class="summary-fact-value">${(delta_T / (R_pipe_wall + R_soil_hs)).toFixed(2)} BTU/hr·ft</span>
              </div>
              <div class="summary-fact">
                  <span class="summary-fact-label">Heat Loss Reduction due to Insulation:</span>
                  <span class="summary-fact-value">${(((delta_T / (R_pipe_wall + R_soil_hs)) - Q) / (delta_T / (R_pipe_wall + R_soil_hs)) * 100).toFixed(1)}%</span>
              </div>
          </div>`;
      }

      resultsContainer.innerHTML = resultsHTML;
      resultsContainer.classList.remove('hidden');

      // --- Generate LaTeX Report ---
      generateAndDisplayLatexReport(data, lastCalculationResults);
      

    } catch(e) {
      console.error(e);
      resultsContainer.innerHTML = `<div class="calculation-error"><h3>Calculation Failed</h3><p>An error occurred during calculation. Please check your inputs and ensure all values are valid numbers. The console may have more details.</p><p>${(e as Error).message}</p></div>`;
      resultsContainer.classList.remove('hidden');
      lastCalculationResults = null;
      document.getElementById('latexContainer')?.classList.add('hidden');
      document.getElementById('latexPlaceholder')?.classList.remove('hidden');
    }
  });

  const generateAndDisplayLatexReport = (data: ReturnType<typeof getFormData>, results: any) => {
    const latexContainer = document.getElementById('latexContainer');
    const latexPlaceholder = document.getElementById('latexPlaceholder');
    const latexOutputEl = document.getElementById('latexReportOutput');
    if (!latexContainer || !latexPlaceholder || !latexOutputEl) return;
  
    const escapeLatex = (str: string | undefined | null) => {
      if (!str) return 'N/A';
      return str.replace(/&/g, '\\&').replace(/%/g, '\\%').replace(/\$/g, '\\$').replace(/#/g, '\\#').replace(/_/g, '\\_').replace(/\{/g, '\\{').replace(/\}/g, '\\}').replace(/~/g, '\\textasciitilde{}').replace(/\^/g, '\\textasciicircum{}').replace(/\\/g, '\\textbackslash{}');
    };
  
    const formatValue = (value: any, unit = '') => {
      const valStr = String(value).trim();
      if (valStr === '' || valStr === 'N/A' || valStr === 'null' || valStr === 'undefined') {
        return 'N/A';
      }
      return `${escapeLatex(valStr)}${unit ? ` ${unit}` : ''}`;
    };

    const formatLongText = (text: string | undefined | null) => {
        if (!text) return 'N/A';
        return `\\\\ \\parbox[t]{\\dimexpr\\linewidth-9em}{${escapeLatex(text)}}`;
    }
  
    const formatItem = (label: string, value: string | undefined | null) => `\\item[\\textbf{${label}:}] ${value ?? 'N/A'}`;
    const beginList = '\\begin{description}[font=\\normalfont, style=unboxed, leftmargin=0pt]';
    const endList = '\\end{description}';
  
    let latexString = `\\documentclass{article}
\\usepackage[utf8]{inputenc}
\\usepackage{geometry}
\\usepackage{enumitem}
\\geometry{a4paper, margin=1in}

\\begin{document}

\\title{Thermal–Gas Line Encroachment Assessment Report}
\\author{${escapeLatex(data.engineerName || 'N/A')}}
\\date{${escapeLatex(data.date || 'N/A')}}
\\maketitle

% --- Evaluation Information ---
\\section*{Evaluation Information}
${beginList}
  ${formatItem('Date', formatValue(data.date))}
  ${formatItem(`Evaluator Name${(data.evaluatorNames || []).length > 1 ? 's' : ''}`, formatValue(data.evaluatorNames.join(', ')))}
  ${formatItem('Engineer’s Name', formatValue(data.engineerName))}
  ${formatItem('Project Name', formatValue(data.projectName))}
  ${formatItem('Project Location', formatValue(data.projectLocation))}
  ${formatItem('Project Description', formatLongText(data.projectDescription))}
${endList}

% --- Heat Source Data ---
\\section*{Heat Source Data}
${beginList}`;
    if (!data.isHeatSourceApplicable) {
        latexString += `  ${formatItem('Heat Source Status', 'No applicable heat source type selected.')}`;
    } else {
        latexString += `
  ${formatItem('Heat Source Type', formatValue(data.heatSourceType))}
  ${formatItem('Operator Contact', formatValue(data.operatorName))}
  ${formatItem('Operator Company', formatValue(data.operatorCompanyName))}
  ${formatItem('Operator Address', formatValue(data.operatorCompanyAddress))}
  ${formatItem('Operator Contact Info', formatValue(data.operatorContactInfo))}
  ${formatItem('Registered with 811', formatValue(data.isRegistered811) + ` (${data.registered811Confirmation})`)}
  ${formatItem('Data Confirmation Date', formatValue(data.confirmationDate))}
  ${formatItem('Max Operating Temp', formatValue(data.maxTemp, '°F') + ` (${data.tempType})`)}
  ${formatItem('Max Operating Pressure', formatValue(data.maxPressure, 'psig') + ` (${data.pressureType})`)}
  ${formatItem('Line Age', formatValue(data.heatSourceAge, 'years') + ` (${data.ageType})`)}
  ${formatItem('System Duty Cycle', formatLongText(data.systemDutyCycle) + ` (${data.systemDutyCycleType})`)}
  ${formatItem('Pipe Casing Info', formatLongText(data.pipeCasingInfo) + ` (${data.pipeCasingInfoType})`)}
  ${formatItem('Evidence of Surface Heat Loss', formatLongText(data.heatLossEvidence) + ` (Source: ${data.heatLossEvidenceSource})`)}
  ${formatItem('Pipeline Nominal Diameter', formatValue(data.pipelineDiameter, 'inches') + ` (${data.diameterType})`)}
  ${formatItem('Pipe Material', formatValue(data.heatSourcePipeMaterial) + ` (${data.materialType})`)}
  ${formatItem('Pipe Wall Thickness', formatValue(data.heatSourceWallThickness, 'inches') + ` (${data.wallThicknessType})`)}
  ${formatItem('Connection Types', formatLongText(data.connectionsValue.join(', ')) + ` (${data.connectionTypesType})`)}
  ${formatItem('Pipe Insulation Type', formatValue(data.insulationType) + ` (${data.insulationTypeType})`)}
  ${data.insulationType.startsWith('Other') ? `${formatItem('Custom Insulation K', formatValue(data.customInsulationThermalConductivity, 'BTU/hr·ft·°F'))}` : ''}
  ${data.insulationType !== 'None' ? `${formatItem('Insulation Thickness', formatValue(data.insulationThickness, 'inches') + ` (${data.insulationThicknessType})`)}` : ''}
  ${data.insulationType !== 'None' ? `${formatItem('Insulation Condition', formatLongText(data.insulationCondition) + ` (Source: ${data.insulationConditionSource})`)}` : ''}
  ${formatItem('Depth to Centerline', formatValue(data.heatSourceDepth, 'feet') + ` (${data.depthType})`)}
  ${formatItem('Additional Operator Info', formatLongText(data.additionalInfo))}`;
    }
latexString += `
${endList}

% --- Gas Line Data ---
\\section*{Gas Line Data}
${beginList}
  ${formatItem('Gas Line Operator Name', formatValue(data.gasOperatorName))}
  ${formatItem('Distribution Operating Co.', formatValue(data.gasDoc))}
  ${formatItem('Max Operating Pressure', formatValue(data.gasMaxPressure, 'psig'))}
  ${formatItem('Installation Year', formatValue(data.gasInstallationYear))}
  ${formatItem('Pipeline Nominal Diameter', formatValue(data.gasPipelineDiameter, 'inches'))}
  ${formatItem('Pipe Sizing Standard', formatValue(data.gasPipeSizingStandard.toUpperCase()))}
  ${formatItem('Gas Pipeline Material', formatValue(data.gasPipeMaterial))}
  ${data.gasPipeMaterial.startsWith('Other') ? `${formatItem('Custom Material K', formatValue(data.customGasThermalConductivity, 'BTU/hr·ft·°F'))}` : ''}
  ${data.gasPipeWallThickness !== 'N/A' ? `${formatItem('Gas Pipe Wall Thickness', formatValue(data.gasPipeWallThickness, 'inches'))}` : ''}
  ${data.gasPipeSDR !== 'N/A' ? `${formatItem('Gas Pipe SDR', formatValue(data.gasPipeSDR))}` : ''}`;
  if (['hdpe', 'mdpe', 'aldyl'].includes(data.gasPipeMaterialValue)) {
    latexString += `
  ${formatItem('Material Continuous Temp Limit', formatValue(data.gasPipeContinuousLimit, '°F'))}
  ${formatItem('Common Utility Cap Temp', formatValue(data.gasPipeUtilityCap, '°F'))}
  ${formatItem('Operational Temp Limit', `\\textbf{${formatValue('70', '°F')} (Max for all plastic pipes)}`)}
  ${formatItem('Material Notes', formatLongText(data.gasPipeNotes))}`;
  }
latexString += `
  ${data.gasCoatingType !== 'N/A' ? `${formatItem('Coating Type', formatValue(data.gasCoatingType))}` : ''}
  ${(data.gasCoatingMaxTemp !== 'N/A' && data.gasCoatingMaxTemp !== 'Custom') ? `${formatItem('Coating Max Allowable Temp', formatValue(data.gasCoatingMaxTemp, '°F'))}` : ''}
  ${formatItem('Orientation to Heat Source', formatValue(data.gasLineOrientation))}
  ${formatItem('Depth to Centerline', formatValue(data.depthOfBurialGasLine, 'feet'))}
  ${data.parallelDistance !== 'N/A' ? `${formatItem('Distance Between Lines', formatValue(data.parallelDistance, 'feet'))}` : ''}
  ${data.parallelLength !== 'N/A' ? `${formatItem('Length of Parallel Section', formatValue(data.parallelLength, 'feet'))}` : ''}
  ${data.latitude !== 'N/A' ? `${formatItem('Intersection Latitude', formatValue(data.latitude))}` : ''}
  ${data.longitude !== 'N/A' ? `${formatItem('Intersection Longitude', formatValue(data.longitude))}` : ''}
${endList}

% --- Soil Data ---
\\section*{Soil \\& Bedding Data}
${beginList}
  ${formatItem('Native Soil Type', formatValue(data.soilType))}
  ${formatItem('Native Soil Thermal Conductivity', formatValue(data.soilThermalConductivity, 'BTU/hr·ft·°F'))}
  ${formatItem('Soil Moisture Content', formatValue(data.soilMoistureContent, '%'))}
  ${formatItem('Average Ground Surface Temp', formatValue(data.groundSurfaceTemperature, '°F'))}
  ${formatItem('Evidence of Water Infiltration', formatValue(data.waterInfiltration))}
  ${formatItem('Infiltration Comments', formatLongText(data.waterInfiltrationComments))}
  ${formatItem('Heat Source Bedding Type', formatValue(data.heatSourceBeddingType))}
  ${data.heatSourceBeddingType === 'sand' ? `${formatItem('HS Bedding Dims (T/B/L/R)', `${formatValue(data.heatSourceBeddingTop, '"')} / ${formatValue(data.heatSourceBeddingBottom, '"')} / ${formatValue(data.heatSourceBeddingLeft, '"')} / ${formatValue(data.heatSourceBeddingRight, '"')}`)}` : ''}
  ${data.heatSourceBeddingType === 'sand' ? `${formatItem('HS Bedding K', data.heatSourceBeddingUseCustomK === 'yes' ? `${formatValue(data.heatSourceBeddingCustomK, 'BTU/hr·ft·°F')} (Custom)` : '0.20 BTU/hr·ft·°F (Default)')}` : ''}
  ${formatItem('Gas Line Bedding Type', formatValue(data.gasLineBeddingType))}
  ${data.gasLineBeddingType === 'sand' ? `${formatItem('Gas Bedding Dims (T/B/L/R)', `${formatValue(data.gasBeddingTop, '"')} / ${formatValue(data.gasBeddingBottom, '"')} / ${formatValue(data.gasBeddingLeft, '"')} / ${formatValue(data.gasBeddingRight, '"')}`)}` : ''}
  ${data.gasLineBeddingType === 'sand' ? `${formatItem('Gas Bedding K', data.gasBeddingUseCustomK === 'yes' ? `${formatValue(data.gasBeddingCustomK, 'BTU/hr·ft·°F')} (Custom)` : '0.20 BTU/hr·ft·°F (Default)')}` : ''}
${endList}

% --- Field Visit ---
\\section*{Field Visit}
${beginList}
    ${formatItem('Date of Visit', formatValue(data.visitDate))}
    ${formatItem('Personnel on Site', formatLongText(data.sitePersonnel))}
    ${formatItem('Weather and Site Conditions', formatLongText(data.siteConditions))}
    ${formatItem('Field Observations', formatLongText(data.fieldObservations))}
${endList}

% --- Calculation Results ---
\\section*{Calculation Results}
${beginList}
  ${formatItem('Calculated Temp at Gas Line Surface', `\\textbf{${formatValue(results['Calculated Temperature at Gas Line Surface (°F)'], '°F')}}`)}
  ${formatItem('Temperature Rise Above Ground', formatValue(results['Temperature Rise Above Ground (°F)'], '°F'))}
  ${formatItem('Heat Flux at Gas Pipe Surface', formatValue(results['Heat Flux at Gas Pipe Surface (BTU/hr·ft²)'], 'BTU/hr$\\cdot$ft$^2$'))}
  ${formatItem('Heat Loss per Unit Length (Q)', formatValue(results['Heat Loss per Unit Length (Q, BTU/hr·ft)'], 'BTU/hr·ft'))}
  ${formatItem('Total Thermal Resistance (R\\_total)', formatValue(results['Total Thermal Resistance (R_total, hr·ft·°F/BTU)'], 'hr·ft·°F/BTU'))}
  ${formatItem('Pipe Wall Thermal Resistance (R\\_pipe\\_wall)', formatValue(results['Pipe Wall Thermal Resistance (hr·ft·°F/BTU)'], 'hr·ft·°F/BTU'))}
  ${formatItem('Insulation Thermal Resistance (R\\_insulation)', formatValue(results['Insulation Thermal Resistance (hr·ft·°F/BTU)'], 'hr·ft·°F/BTU'))}
  ${formatItem('Soil Thermal Resistance (R\\_soil\\_hs)', formatValue(results['Soil Thermal Resistance (hr·ft·°F/BTU)'], 'hr·ft·°F/BTU'))}
${endList}

% --- Recommendations ---
\\section*{Recommendations}
\\begin{enumerate}
  ${data.recommendations.filter(rec => rec.trim() !== '').map(rec => `\\item ${escapeLatex(rec)}`).join('\n  ')}
\\end{enumerate}

\\end{document}
`;
  
    latexOutputEl.textContent = latexString;
    latexContainer.classList.remove('hidden');
    latexPlaceholder.classList.add('hidden');
  };

  const copyLatexButton = document.getElementById('copyLatexButton');
  copyLatexButton?.addEventListener('click', () => {
    const latexCode = document.getElementById('latexReportOutput')?.textContent;
    if (latexCode) {
      navigator.clipboard.writeText(latexCode).then(() => {
        const originalText = copyLatexButton.textContent;
        copyLatexButton.textContent = 'Copied!';
        setTimeout(() => {
          copyLatexButton.textContent = originalText;
        }, 2000);
      }).catch(err => {
        console.error('Failed to copy text: ', err);
        alert('Failed to copy LaTeX code to clipboard.');
      });
    }
  });

});