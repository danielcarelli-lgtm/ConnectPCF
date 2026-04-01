/* eslint-disable @microsoft/power-apps/avoid-dom-form */
import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface IModeContextInfo {
    contextInfo: {
        entityId: string;
        entityTypeName: string;
    };
}

interface IXrmWindow {
    Xrm?: {
        Page?: {
            data?: {
                refresh: (save: boolean) => Promise<void>;
            };
        };
    };
}

export class ConnectPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged: () => void;

    // Elementos de la UI
    private mainWrapper: HTMLDivElement;
    private stepperContainer: HTMLDivElement;
    private contentContainer: HTMLDivElement;
    private versionLabel: HTMLDivElement;

    // Valores de la Razón para el estado
    private readonly STATUS_PENDIENTE = 1;
    private readonly STATUS_EN_CURSO = 919690002;
    private readonly STATUS_COMPLETADO = 2;
    private readonly STATUS_CANCELADO = 919690001;

    // Valores del Estado (StateCode)
    private readonly STATE_ACTIVE = 0;
    private readonly STATE_INACTIVE = 1;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;

        this.mainWrapper = document.createElement("div");
        this.mainWrapper.className = "connect-pcf-wrapper";

        this.stepperContainer = document.createElement("div");
        this.stepperContainer.className = "connect-stepper";

        this.contentContainer = document.createElement("div");
        this.contentContainer.className = "connect-content";

        // Etiqueta de versión
        this.versionLabel = document.createElement("div");
        this.versionLabel.className = "version-label";
        this.versionLabel.innerText = "v0.0.20";

        this.mainWrapper.appendChild(this.stepperContainer);
        this.mainWrapper.appendChild(this.contentContainer);
        this.mainWrapper.appendChild(this.versionLabel);
        this.container.appendChild(this.mainWrapper);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        const currentState = context.parameters.statuscode.raw || this.STATUS_PENDIENTE;

        this.renderStepper(currentState);
        this.renderContent(currentState);
    }

    private renderStepper(currentState: number): void {
        this.stepperContainer.innerHTML = ""; 

        const orderMap: Record<number, number> = {
            [this.STATUS_PENDIENTE]: 1,
            [this.STATUS_EN_CURSO]: 2,
            [this.STATUS_COMPLETADO]: 3,
            [this.STATUS_CANCELADO]: 3 
        };

        const currentOrder = orderMap[currentState] || 0;

        const steps = [
            { id: this.STATUS_PENDIENTE, label: "Pendiente" },
            { id: this.STATUS_EN_CURSO, label: "En Curso" },
            currentState === this.STATUS_CANCELADO 
                ? { id: this.STATUS_CANCELADO, label: "Cancelado" } 
                : { id: this.STATUS_COMPLETADO, label: "Completado" }
        ];

        const flowWrapper = document.createElement("div");
        flowWrapper.className = "flow-wrapper";

        const stepsContainer = document.createElement("div");
        stepsContainer.className = "steps-container";

        steps.forEach((step, index) => {
            const stepEl = document.createElement("div");
            stepEl.className = "stepper-item";
            
            const stepOrder = orderMap[step.id];

            if (currentState === step.id) {
                stepEl.classList.add("active");
            } else if (currentOrder > stepOrder) {
                stepEl.classList.add("completed");
            }

            const stepCircle = document.createElement("div");
            stepCircle.className = "step-circle";
            stepCircle.innerText = (index + 1).toString();

            const stepLabel = document.createElement("div");
            stepLabel.className = "step-label";
            stepLabel.innerText = step.label;

            stepEl.appendChild(stepCircle);
            stepEl.appendChild(stepLabel);
            stepsContainer.appendChild(stepEl);

            if (index < steps.length - 1) {
                const line = document.createElement("div");
                line.className = "step-line";
                if (currentOrder > stepOrder) {
                    line.classList.add("completed-line");
                }
                stepsContainer.appendChild(line);
            }
        });

        const graphicDiv = document.createElement("div");
        graphicDiv.className = "state-graphic";
        
        let graphicContent = "🏁"; 
        if (currentState === this.STATUS_PENDIENTE) graphicContent = "🏁🏃‍♂️"; 
        else if (currentState === this.STATUS_EN_CURSO) graphicContent = "👨‍💻⚙️"; 
        else if (currentState === this.STATUS_COMPLETADO) graphicContent = "🎉🏆"; 
        else if (currentState === this.STATUS_CANCELADO) graphicContent = "😢🚫"; 

        graphicDiv.innerText = graphicContent;

        flowWrapper.appendChild(stepsContainer);
        flowWrapper.appendChild(graphicDiv);
        this.stepperContainer.appendChild(flowWrapper);
    }

    private renderContent(currentState: number): void {
        this.contentContainer.innerHTML = ""; 

        if (currentState === this.STATUS_PENDIENTE) {
            this.renderEtapa1();
        } else if (currentState === this.STATUS_EN_CURSO) {
            this.renderEtapa2();
        } else if (currentState === this.STATUS_COMPLETADO) {
            this.renderEtapa3();
        } else if (currentState === this.STATUS_CANCELADO) {
            this.renderEtapa4();
        } else {
            const defaultMsg = document.createElement("p");
            defaultMsg.innerText = "Estado actual no configurado en el cuadro de mando.";
            this.contentContainer.appendChild(defaultMsg);
        }
    }

    private renderEtapa1(): void {
        const messagePara = document.createElement("p");
        messagePara.className = "stage-message";
        
        const highlightSpan = document.createElement("span");
        highlightSpan.className = "status-highlight";
        highlightSpan.innerText = "Pendiente";

        messagePara.appendChild(document.createTextNode("Este Connect se encuentra en estado "));
        messagePara.appendChild(highlightSpan);
        messagePara.appendChild(document.createTextNode(", presiona el botón Comenzar para asignártelo y comenzar a trabajar."));

        const btnGroup = document.createElement("div");
        btnGroup.className = "button-group";

        const btnComenzar = document.createElement("button");
        btnComenzar.className = "btn-modern btn-primary";
        btnComenzar.innerText = "Comenzar";
        btnComenzar.onclick = this.onComenzarClick.bind(this);

        const btnCancelar = document.createElement("button");
        btnCancelar.className = "btn-modern btn-danger";
        btnCancelar.innerText = "Cancelar";
        btnCancelar.onclick = this.onCancelarClick.bind(this);

        btnGroup.appendChild(btnComenzar);
        btnGroup.appendChild(btnCancelar);

        this.contentContainer.appendChild(messagePara);
        this.contentContainer.appendChild(btnGroup);
    }

    private renderEtapa2(): void {
        const cuentaData = this.context.parameters.sec_cuentaid.raw;
        const ubicacionData = this.context.parameters.sec_ubicaciontecnicaid.raw;

        const hasCuenta = cuentaData != null && cuentaData.length > 0;
        const hasUbicacion = ubicacionData != null && ubicacionData.length > 0;

        const cuentaName = hasCuenta ? cuentaData![0].name : "";
        const ubicacionName = hasUbicacion ? ubicacionData![0].name : "";

        const listContainer = document.createElement("ul");
        listContainer.className = "task-list";

        const task1 = document.createElement("li");
        task1.className = "task-item";
        const icon1 = this.createStatusIcon(hasCuenta);
        const text1 = document.createElement("span");
        
        if (hasCuenta) {
            text1.innerText = `Ya hemos agregado la cuenta (${cuentaName}).`;
            text1.className = "text-orange-matte";
        } else {
            text1.innerText = "1 - Crear nueva cuenta en CEP o asociar una existente en la pestaña “Vinculado a”.";
        }
        task1.appendChild(icon1);
        task1.appendChild(text1);

        const task2 = document.createElement("li");
        task2.className = "task-item";
        const icon2 = this.createStatusIcon(hasUbicacion);
        const text2 = document.createElement("span");
        
        if (hasUbicacion) {
            text2.innerText = `Ya hemos agregado la ubicación (${ubicacionName}).`;
            text2.className = "text-orange-matte";
        } else {
            text2.innerText = "2 - Crear un centro de trabajo (Ubicación o Site) o asociar una existente en la pestaña “Vinculado a”.";
        }
        task2.appendChild(icon2);
        task2.appendChild(text2);

        listContainer.appendChild(task1);
        listContainer.appendChild(task2);

        const btnGroup = document.createElement("div");
        btnGroup.className = "button-group";

        const btnCrearOT = document.createElement("button");
        btnCrearOT.className = "btn-modern btn-success";
        btnCrearOT.innerText = "Crear Orden de Trabajo";
        if (!hasCuenta || !hasUbicacion) {
            btnCrearOT.disabled = true;
            btnCrearOT.title = "Complete los campos de cuenta y ubicación técnica para habilitar.";
        } else {
            btnCrearOT.onclick = this.onCrearOTClick.bind(this);
        }

        const btnCancelar = document.createElement("button");
        btnCancelar.className = "btn-modern btn-danger";
        btnCancelar.innerText = "Cancelar";
        btnCancelar.onclick = this.onCancelarClick.bind(this);

        const btnRefresh = document.createElement("button");
        btnRefresh.className = "btn-modern btn-secondary";
        btnRefresh.innerText = "Refrescar"; 
        btnRefresh.onclick = () => {
            this.updateView(this.context);
        };

        btnGroup.appendChild(btnCrearOT);
        btnGroup.appendChild(btnCancelar);
        btnGroup.appendChild(btnRefresh);

        this.contentContainer.appendChild(listContainer);
        this.contentContainer.appendChild(btnGroup);
    }

    private renderEtapa3(): void {
        const messagePara = document.createElement("p");
        messagePara.className = "stage-message";
        
        const highlightSpan = document.createElement("span");
        highlightSpan.className = "status-highlight success-highlight";
        highlightSpan.innerText = "Enhorabuena, hemos completado el proceso";

        messagePara.appendChild(highlightSpan);
        this.contentContainer.appendChild(messagePara);
    }

    private renderEtapa4(): void {
        const messagePara = document.createElement("p");
        messagePara.className = "stage-message";
        
        const highlightSpan = document.createElement("span");
        highlightSpan.className = "status-highlight danger-highlight";
        highlightSpan.innerText = "Vaya.. este proceso ha sido cancelado";

        messagePara.appendChild(highlightSpan);
        this.contentContainer.appendChild(messagePara);
    }

    private createStatusIcon(isValid: boolean): HTMLSpanElement {
        const icon = document.createElement("span");
        icon.className = isValid ? "icon-check" : "icon-cross";
        icon.innerText = isValid ? "✔" : "✖";
        return icon;
    }

    private refreshDynamicsForm(): void {
        const globalWindow = window as unknown as IXrmWindow;
        if (globalWindow.Xrm?.Page?.data) {
            globalWindow.Xrm.Page.data.refresh(true).catch((err: unknown) => {
                console.warn("No se pudo refrescar el formulario automáticamente:", err);
            });
        }
    }

    private showErrorModal(errorText: string): void {
        const modal = document.createElement("div");
        modal.className = "error-modal-overlay";
        
        const modalContent = document.createElement("div");
        modalContent.className = "error-modal-content";
        
        const title = document.createElement("h3");
        title.innerText = "⚠️ Error al crear Orden de Trabajo";
        title.style.color = "#b75d5d";
        title.style.margin = "0 0 10px 0";
        
        const instructions = document.createElement("p");
        instructions.innerText = "Por favor, copia todo el texto de abajo para reportar el problema exacto de Dataverse:";
        instructions.style.fontSize = "12px";
        instructions.style.marginBottom = "10px";
        
        const textArea = document.createElement("textarea");
        textArea.readOnly = true;
        textArea.value = errorText;
        textArea.className = "error-textarea";
        
        const btnGroup = document.createElement("div");
        btnGroup.style.display = "flex";
        btnGroup.style.justifyContent = "flex-end";
        btnGroup.style.marginTop = "15px";

        const closeBtn = document.createElement("button");
        closeBtn.className = "btn-modern btn-secondary";
        closeBtn.innerText = "Cerrar";
        closeBtn.onclick = () => {
            if (modal.parentElement) {
                modal.parentElement.removeChild(modal);
            }
        };
        
        btnGroup.appendChild(closeBtn);

        modalContent.appendChild(title);
        modalContent.appendChild(instructions);
        modalContent.appendChild(textArea);
        modalContent.appendChild(btnGroup);
        modal.appendChild(modalContent);
        
        this.mainWrapper.appendChild(modal);
    }

    private async onComenzarClick(): Promise<void> {
        try {
            const modeInfo = this.context.mode as unknown as IModeContextInfo;
            const recordId = modeInfo.contextInfo.entityId;
            const entityName = modeInfo.contextInfo.entityTypeName;
            const currentUserId = this.context.userSettings.userId;

            if (!recordId) {
                alert("Por favor, guarde el registro antes de comenzar.");
                return;
            }

            const data: Record<string, string | number> = {
                "statecode": this.STATE_ACTIVE,
                "statuscode": this.STATUS_EN_CURSO,
                "ownerid@odata.bind": `/systemusers(${currentUserId.replace("{", "").replace("}", "")})`
            };

            await this.context.webAPI.updateRecord(entityName, recordId, data);
            
            this.notifyOutputChanged();
            this.refreshDynamicsForm();

        } catch (error: unknown) {
            console.error("Error al actualizar el registro:", error);
            const errorMessage = error instanceof Error ? error.message : String(error);
            alert("Hubo un error al intentar comenzar: " + errorMessage);
        }
    }

    private async onCancelarClick(): Promise<void> {
        try {
            const modeInfo = this.context.mode as unknown as IModeContextInfo;
            const recordId = modeInfo.contextInfo.entityId;
            const entityName = modeInfo.contextInfo.entityTypeName;

            if (!recordId) {
                alert("Por favor, guarde el registro antes de cancelar.");
                return;
            }

            const data: Record<string, string | number> = {
                "statecode": this.STATE_INACTIVE,
                "statuscode": this.STATUS_CANCELADO
            };

            await this.context.webAPI.updateRecord(entityName, recordId, data);
            
            this.notifyOutputChanged();
            this.refreshDynamicsForm();

        } catch (error: unknown) {
            console.error("Error al cancelar el registro:", error);
            const errorMessage = error instanceof Error ? error.message : String(error);
            alert("Hubo un error al intentar cancelar: " + errorMessage);
        }
    }

    private async onCrearOTClick(): Promise<void> {
        let payloadEnviado = "";

        try {
            const modeInfo = this.context.mode as unknown as IModeContextInfo;
            const recordId = modeInfo.contextInfo.entityId;
            const entityName = modeInfo.contextInfo.entityTypeName;

            if (!recordId) {
                alert("Por favor, guarde el registro antes de continuar.");
                return;
            }

            const cuentaData = this.context.parameters.sec_cuentaid.raw;
            const ubicacionData = this.context.parameters.sec_ubicaciontecnicaid.raw;

            if (!cuentaData || cuentaData.length === 0 || !ubicacionData || ubicacionData.length === 0) {
                alert("Faltan datos de Cuenta o Ubicación Técnica.");
                return;
            }

            const cuentaId = cuentaData[0].id.replace("{", "").replace("}", "");
            const ubicacionId = ubicacionData[0].id.replace("{", "").replace("}", "");
            
            const tipoIncidenteId = "c7b290d7-1c2d-f111-88b4-7c1e527828ba";
            const prioridadId = "db54a64a-dff5-ed11-8e4b-002248a6ca1f";

            const woData: Record<string, string | number> = {
                "msdyn_serviceaccount@odata.bind": `/accounts(${cuentaId})`,
                "msdyn_FunctionalLocation@odata.bind": `/msdyn_functionallocations(${ubicacionId})`,
                "msdyn_primaryincidenttype@odata.bind": `/msdyn_incidenttypes(${tipoIncidenteId})`,
                "msdyn_priority@odata.bind": `/msdyn_priorities(${prioridadId})`,
                "msdyn_workordersummary": "Generado desde proceso Connect"
            };

            payloadEnviado = JSON.stringify(woData, null, 2);

            // 1. Creación de Orden de Trabajo
            await this.context.webAPI.createRecord("msdyn_workorder", woData);

            // 2. Actualización de Connect a Completado (State=1, Status=2)
            const updateData: Record<string, string | number> = {
                "statecode": this.STATE_INACTIVE, 
                "statuscode": this.STATUS_COMPLETADO 
            };
            await this.context.webAPI.updateRecord(entityName, recordId, updateData);
            
            this.notifyOutputChanged();
            this.refreshDynamicsForm();

        } catch (error: unknown) {
            console.error("Error completo al crear Orden de Trabajo:", error);
            
            let errorDetails = "";
            if (error instanceof Error) {
                errorDetails = error.message;
            } else if (typeof error === "object" && error !== null) {
                const errObj = error as Record<string, unknown>;
                if ('message' in errObj && typeof errObj.message === 'string') {
                    errorDetails = errObj.message;
                }
                try {
                    errorDetails += "\n\n--- JSON RAW DEL ERROR ---\n" + JSON.stringify(error, null, 2);
                } catch (e) {
                    errorDetails += "\n\n(No se pudo convertir a JSON el objeto de error)";
                }
            } else {
                errorDetails = String(error);
            }

            const mensajeFinal = `--- PAYLOAD ENVIADO ---\n${payloadEnviado}\n\n--- DETALLE DEL ERROR ---\n${errorDetails}`;
            this.showErrorModal(mensajeFinal);
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.container.innerHTML = "";
    }
}