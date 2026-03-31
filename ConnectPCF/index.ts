import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface IModeContextInfo {
    contextInfo: {
        entityId: string;
        entityTypeName: string;
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
        this.versionLabel.innerText = "v0.0.8";

        this.mainWrapper.appendChild(this.stepperContainer);
        this.mainWrapper.appendChild(this.contentContainer);
        this.mainWrapper.appendChild(this.versionLabel);
        this.container.appendChild(this.mainWrapper);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        // Al ser un OptionSet, el .raw seguirá devolviendo el valor numérico de la opción elegida
        const currentState = context.parameters.statuscode.raw || this.STATUS_PENDIENTE;

        this.renderStepper(currentState);
        this.renderContent(currentState);
    }

    private renderStepper(currentState: number): void {
        this.stepperContainer.innerHTML = ""; 

        const steps = [
            { id: this.STATUS_PENDIENTE, label: "Pendiente" },
            { id: this.STATUS_EN_CURSO, label: "En Curso" }
        ];

        steps.forEach((step, index) => {
            const stepEl = document.createElement("div");
            stepEl.className = "stepper-item";
            
            if (currentState === step.id) {
                stepEl.classList.add("active");
            } else if (currentState > step.id) {
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
            this.stepperContainer.appendChild(stepEl);

            if (index < steps.length - 1) {
                const line = document.createElement("div");
                line.className = "step-line";
                if (currentState > step.id) {
                    line.classList.add("completed-line");
                }
                this.stepperContainer.appendChild(line);
            }
        });
    }

    private renderContent(currentState: number): void {
        this.contentContainer.innerHTML = ""; 

        if (currentState === this.STATUS_PENDIENTE) {
            this.renderEtapa1();
        } else if (currentState === this.STATUS_EN_CURSO) {
            this.renderEtapa2();
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

        const btnComenzar = document.createElement("button");
        btnComenzar.className = "btn-modern btn-primary";
        btnComenzar.innerText = "Comenzar";
        btnComenzar.onclick = this.onComenzarClick.bind(this);

        this.contentContainer.appendChild(messagePara);
        this.contentContainer.appendChild(btnComenzar);
    }

    private renderEtapa2(): void {
        const cuentaData = this.context.parameters.sec_cuentaid.raw;
        const ubicacionData = this.context.parameters.sec_ubicaciontecnicaid.raw;

        const hasCuenta = cuentaData != null && cuentaData.length > 0;
        const hasUbicacion = ubicacionData != null && ubicacionData.length > 0;

        const listContainer = document.createElement("ul");
        listContainer.className = "task-list";

        const task1 = document.createElement("li");
        task1.className = "task-item";
        const icon1 = this.createStatusIcon(hasCuenta);
        const text1 = document.createElement("span");
        text1.innerText = "1 - Crear nueva cuenta en CEP o asociar una existente en la pestaña “Vinculado a”.";
        task1.appendChild(icon1);
        task1.appendChild(text1);

        const task2 = document.createElement("li");
        task2.className = "task-item";
        const icon2 = this.createStatusIcon(hasUbicacion);
        const text2 = document.createElement("span");
        text2.innerText = "2 - Crear un centro de trabajo (Ubicación o Site) o asociar una existente en la pestaña “Vinculado a”.";
        task2.appendChild(icon2);
        task2.appendChild(text2);

        listContainer.appendChild(task1);
        listContainer.appendChild(task2);

        const btnRefresh = document.createElement("button");
        btnRefresh.className = "btn-icon-refresh";
        btnRefresh.innerHTML = "🔄 Refrescar"; 
        btnRefresh.onclick = () => {
            this.updateView(this.context);
        };

        this.contentContainer.appendChild(listContainer);
        this.contentContainer.appendChild(btnRefresh);
    }

    private createStatusIcon(isValid: boolean): HTMLSpanElement {
        const icon = document.createElement("span");
        icon.className = isValid ? "icon-check" : "icon-cross";
        icon.innerText = isValid ? "✔" : "✖";
        return icon;
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
                "statuscode": this.STATUS_EN_CURSO,
                "ownerid@odata.bind": `/systemusers(${currentUserId.replace("{", "").replace("}", "")})`
            };

            await this.context.webAPI.updateRecord(entityName, recordId, data);
            
            this.notifyOutputChanged();

        } catch (error: unknown) {
            console.error("Error al actualizar el registro:", error);
            const errorMessage = error instanceof Error ? error.message : String(error);
            alert("Hubo un error al intentar comenzar: " + errorMessage);
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.container.innerHTML = "";
    }
}