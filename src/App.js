import logo from './logo.svg';

import React, { useState } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import "./App.css";
import businessData from "./data/businessData"; 

const App = () => {

  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const handleLogin = () => {
    instance
      .loginPopup({
        scopes: ["User.Read"], // Permisos configurados en Azure AD
      })
      .then((response) => {
        console.log("Logged in:", response);
      })
      .catch((error) => {
        console.error("Login failed:", error);
      });
  };

  const handleLogout = () => {
    instance.logoutPopup();
  };

  const [data] = useState(businessData);

  const [selectedBusinessArea, setSelectedBusinessArea] = useState("");
  const [selectedDomain, setSelectedDomain] = useState("");
  const [selectedService, setSelectedService] = useState("");
  const [selectedMacroCapability, setSelectedMacroCapability] = useState("");
  const [selectedCapabilityLevel1, setSelectedCapabilityLevel1] = useState("");
  const [apiType, setApiType] = useState("");
  const [apiVersion, setApiVersion] = useState("");
  const [method, setMethod] = useState("GET");
  const [actionTerm, setActionTerm] = useState("");
  const [behaviorQualifier, setBehaviorQualifier] = useState("");
  const [generatedData, setGeneratedData] = useState({});
  const [code4Project, setCode4Project] = useState("");

  const serviceDomainValue =
    selectedDomain && selectedCapabilityLevel1
      ? `${selectedCapabilityLevel1.toLowerCase().replace(/\s+/g, "-")}-${selectedDomain.toLowerCase().replace(/\s+/g, "-")}`
      : "";

  const pbcValue =
    selectedCapabilityLevel1 && selectedBusinessArea && selectedDomain
      ? `${selectedCapabilityLevel1.toLowerCase().replace(/\s+/g, "-")}-${selectedBusinessArea.toLowerCase().replace(/\s+/g, "-")}-${selectedDomain.toLowerCase().replace(/\s+/g, "-")}`
      : "";

  const actionTermsByMethod = {
    GET: ["retrieve", "notify"],
    POST: ["create", "initiate", "register", "evaluate", "provide", "redemption"],
  };

  const availableActionTerms = actionTermsByMethod[method] || [];

  const handleGenerateAPI = () => {
    const apiBasePath = `/${serviceDomainValue?.toLowerCase().replace(/\s/g, "-")}/v${apiVersion}`;
    const apiPathOperation = `/${behaviorQualifier?.toLowerCase().replace(/\s/g, "-")}/${actionTerm?.toLowerCase()}`;
    const apiURI = `${apiBasePath}${apiPathOperation}`;
    const objectRequest = `${actionTerm?.charAt(0).toUpperCase() + actionTerm?.slice(1)}${behaviorQualifier
      ?.split(" ")
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join("")}`;
    const objectResponse = `${objectRequest}Response`;
    const behaviorQua = `${behaviorQualifier?.toLowerCase().replace(/\s/g, "-")}`;
    const repositoryValue = 
      apiType && code4Project && serviceDomainValue && behaviorQua
        ? `${apiType.toLowerCase().replace(/\s/g, "-")}-${code4Project.toLowerCase().replace(/\s/g, "-")}-${serviceDomainValue.toLowerCase().replace(/\s/g, "-")}-${behaviorQua.toLowerCase().replace(/\s/g, "-")}`
        : "";

    setGeneratedData({
      "API Type": apiType,
      "Business Area": selectedDomain,
      "API Name": serviceDomainValue,
      "Macro Capability": selectedMacroCapability,
      "Capability Level 1": selectedCapabilityLevel1,
      "PBC": pbcValue,
      "API Base Path": apiBasePath,
      "Behavior Qualifier": behaviorQua,
      "API Path Operation": apiPathOperation,
      "API URI Niubiz": apiURI,
      "Object Request": objectRequest,
      "Object Response": objectResponse,
      "Repository Bitbucket": repositoryValue,
    });

    console.log("Generated Data:", {
      "API Type": apiType,
      "Business Area": selectedDomain,
      "API Name": serviceDomainValue,
      "Macro Capability": selectedMacroCapability,
      "Capability Level 1": selectedCapabilityLevel1,
      "PBC": pbcValue, 
      "API Base Path": apiBasePath,
      "Behavior Qualifier": behaviorQua,
      "API Path Operation": apiPathOperation,
      "API URI Niubiz": apiURI,
      "Object Request": objectRequest,
      "Object Response": objectResponse,
      "Repository Bitbucket": repositoryValue,
    });
  };

  const handleDownloadYAML = () => {
    console.log(generatedData)
    if (!Object.keys(generatedData).length) {
      alert("Por favor, genera la API primero antes de descargar el YAML.");
      return;
    }
  
    const yamlContent = `
  openapi: 3.0.1
  info:
    title: ${generatedData["API Name"]}
    description: |
      # PBC: ${generatedData["PBC"]} #
      Semantic API Specification for ${generatedData["API Name"]}
    version: '${apiVersion}'
  tags:
    - name: ${generatedData["Behavior Qualifier"]}
      description: behavior ${generatedData["Behavior Qualifier"]}
  servers:
    - url: 'https://api.${code4Project}.niubiz.com.pe' 
  paths:
    ${generatedData["API URI Niubiz"]}:
      ${method.toLowerCase()}:
        tags:
          - ${generatedData["Behavior Qualifier"]}
        summary: ${actionTerm} ${behaviorQualifier}
        responses:
          '200':
            description: Successful response
            content:
              application/json:
                schema:
                  type: object
                  properties:
                    request:
                      $ref: '#/components/schemas/${generatedData["Object Request"]}'
                    response:
                      $ref: '#/components/schemas/${generatedData["Object Response"]}'
  components:
    schemas:
      ${generatedData["Object Request"]}:
        type: object
        properties:
          exampleProperty:
            type: string
            description: Example property for request
      ${generatedData["Object Response"]}:
        type: object
        properties:
          exampleProperty:
            type: string
            description: Example property for response
    `;
  
    // Crear un blob y descargar el archivo
    const blob = new Blob([yamlContent], { type: "application/x-yaml" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${generatedData["API Name"]}-swagger.yaml`;
    link.click();
  };

  return (
    <div className="container">
      <h1 className="title">Semantic PBCs - APIs - Generator</h1>
      <div className="auth-section">
        {!isAuthenticated ? (
          <button className="auth-button" onClick={handleLogin}>
            Login
          </button>
        ) : (
          <div>
            <button className="logout-button" onClick={handleLogout}>
              Logout
            </button>
          </div>
        )}
      </div>
      {isAuthenticated && (
      <div className="layout">
      <div className="form-container">
      <div className="form-group">
        <label>Business Area:</label>
        <select
          value={selectedBusinessArea}
          onChange={(e) => {
            setSelectedBusinessArea(e.target.value);
            setSelectedDomain("");
            setSelectedService("");
          }}
          className="form-control"
        >
          <option value="">Select Business Area</option>
          {data["business-areas"].map((area) => (
            <option key={area["codigo-ba"]} value={area["business-area"]}>
              {area["business-area"]}
            </option>
          ))}
        </select>
      </div>
      <div className="form-group">
        <label>Business Domain:</label>
        <select
          value={selectedDomain}
          onChange={(e) => {
            setSelectedDomain(e.target.value);
            setSelectedService("");
          }}
          disabled={!selectedBusinessArea}
          className="form-control"
        >
          <option value="">Select Business Domain</option>
          {data["business-areas"]
            .find((area) => area["business-area"] === selectedBusinessArea)
            ?.["business-domains"].map((domain) => (
              <option key={domain["codigo-bd"]} value={domain["business-domain"]}>
                {domain["business-domain"]}
              </option>
            ))}
        </select>
      </div>
      <div className="form-group">
        <label>Macro Capability:</label>
        <select
          value={selectedMacroCapability}
          onChange={(e) => {
            setSelectedMacroCapability(e.target.value);
            setSelectedCapabilityLevel1("");
          }}
          className="form-control"
        >
          <option value="">Select Macro Capability</option>
          {data["macro-capabilitys"].map((macro) => (
            <option key={macro["code-mc"]} value={macro["macro-capability"]}>
              {macro["macro-capability"]}
            </option>
          ))}
        </select>
      </div>
      <div className="form-group">
        <label>Capability Level 1:</label>
        <select
          value={selectedCapabilityLevel1}
          onChange={(e) => setSelectedCapabilityLevel1(e.target.value)}
          disabled={!selectedMacroCapability}
          className="form-control"
        >
          <option value="">Select Capability Level 1</option>
          {data["macro-capabilitys"]
            .find((macro) => macro["macro-capability"] === selectedMacroCapability)
            ?.["capabilities-level-1"].map((capability) => (
              <option key={capability["codigo-cl1"]} value={capability["capability-level-1"]}>
                {capability["capability-level-1"]}
              </option>
            ))}
        </select>
      </div>
      <div className="form-group">
        <label>Service Domain:</label>
        <input
          type="text"
          value={serviceDomainValue}
          disabled
          className="form-control"
        />
      </div>
      <div className="form-group">
        <label>API Type:</label>
        <select
          value={apiType}
          onChange={(e) => setApiType(e.target.value)}
          className="form-control"
        >
          <option value="">Select API Type</option>
          <option value="channel">Channel</option>
          <option value="business">Business</option>
          <option value="support">Support</option>
          <option value="workload">Workload</option>
        </select>
      </div>
      <div className="form-group">
        <label>API Version:</label>
        <select
          value={apiVersion}
          onChange={(e) => setApiVersion(e.target.value)}
          className="form-control"
        >
          <option value="">Select API Version</option>
          <option value="1.0">v1.0</option>
          <option value="2.0">v2.0</option>
          <option value="3.0">v3.0</option>
        </select>
      </div>
      <div className="form-group">
        <label>Method HTTP:</label>
        <select
          value={method}
          onChange={(e) => {
            setMethod(e.target.value);
            setActionTerm("");
          }}
          className="form-control"
        >
          <option value="GET">GET</option>
          <option value="POST">POST</option>
        </select>
      </div>
      <div className="form-group">
        <label>Action Term:</label>
        <select
          value={actionTerm}
          onChange={(e) => setActionTerm(e.target.value)}
          disabled={!method}
          className="form-control"
        >
          <option value="">Select Action Term</option>
          {availableActionTerms.map((term) => (
            <option key={term} value={term}>
              {term}
            </option>
          ))}
        </select>
      </div>
      <div className="form-group">
        <label>Behavior Qualifier:</label>
        <input
          type="text"
          value={behaviorQualifier}
          onChange={(e) => setBehaviorQualifier(e.target.value)}
          className="form-control"
        />
      </div>
      <div className="form-group">
        <label>Code 4 Project:</label>
        <select
          value={code4Project}
          onChange={(e) => setCode4Project(e.target.value)}
          className="form-control"
        >
          <option value="">Select Code 4 Project</option>
          <option value="loyt">Loyalty</option>
          <option value="ryss">Recargas y Servicios</option>
          <option value="ecom">Comercio Electrónico</option>
          <option value="mace">Marca Cerrada</option>
        </select>
      </div>
      <button onClick={handleGenerateAPI} className="generate-button">
        Generate Semantic PBC and API
      </button>
      <button onClick={handleDownloadYAML} className="download-button">
          Descargar API Swagger YAML
        </button>
      </div>

      {/* Columna derecha: Datos generados */}
      {Object.keys(generatedData).length > 0 && (
        <div className="generated-data-container">
          <h2>Semantic PBC-API Generate</h2>
          <ul>
            {Object.entries(generatedData).map(([key, value]) => (
              <li key={key}>
                
                <strong>{key}:</strong> <br/>{value}
              </li>
              
            ))}
          </ul>
        </div>
      )}

    </div>
      )}
    </div>
  );
};

export default App;

