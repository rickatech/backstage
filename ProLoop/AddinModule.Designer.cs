namespace ProLoop
{
    partial class AddinModule
    {
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;
 
        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code
        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.adxWordTaskPanesManager1 = new AddinExpress.WD.ADXWordTaskPanesManager(this.components);
            this.adxWordTaskPanesCollectionItem1 = new AddinExpress.WD.ADXWordTaskPanesCollectionItem(this.components);
            this.adxWordTaskPanesCollectionItem2 = new AddinExpress.WD.ADXWordTaskPanesCollectionItem(this.components);
            this.adxWordTaskPanesCollectionItem3 = new AddinExpress.WD.ADXWordTaskPanesCollectionItem(this.components);
            this.adxWordTaskPanesCollectionItem4 = new AddinExpress.WD.ADXWordTaskPanesCollectionItem(this.components);
            // 
            // adxWordTaskPanesManager1
            // 
            this.adxWordTaskPanesManager1.Items.Add(this.adxWordTaskPanesCollectionItem1);
            this.adxWordTaskPanesManager1.Items.Add(this.adxWordTaskPanesCollectionItem2);
            this.adxWordTaskPanesManager1.Items.Add(this.adxWordTaskPanesCollectionItem3);
            this.adxWordTaskPanesManager1.Items.Add(this.adxWordTaskPanesCollectionItem4);
            this.adxWordTaskPanesManager1.SetOwner(this);
            // 
            // adxWordTaskPanesCollectionItem1
            // 
            this.adxWordTaskPanesCollectionItem1.AlwaysShowHeader = true;
            this.adxWordTaskPanesCollectionItem1.IsHiddenStateAllowed = false;
            this.adxWordTaskPanesCollectionItem1.Position = AddinExpress.WD.ADXWordTaskPanePosition.Right;
            this.adxWordTaskPanesCollectionItem1.RestoreFromMinimizedState = true;
            this.adxWordTaskPanesCollectionItem1.TaskPaneClassName = "ProLoop.Dashboard.ADXWordTaskPaneLogin";
            // 
            // adxWordTaskPanesCollectionItem2
            // 
            this.adxWordTaskPanesCollectionItem2.AlwaysShowHeader = true;
            this.adxWordTaskPanesCollectionItem2.IsHiddenStateAllowed = false;
            this.adxWordTaskPanesCollectionItem2.Position = AddinExpress.WD.ADXWordTaskPanePosition.Right;
            this.adxWordTaskPanesCollectionItem2.RestoreFromMinimizedState = true;
            this.adxWordTaskPanesCollectionItem2.TaskPaneClassName = "ProLoop.Dashboard.ADXWordTaskPaneOpen";
            // 
            // adxWordTaskPanesCollectionItem3
            // 
            this.adxWordTaskPanesCollectionItem3.AlwaysShowHeader = true;
            this.adxWordTaskPanesCollectionItem3.IsHiddenStateAllowed = false;
            this.adxWordTaskPanesCollectionItem3.Position = AddinExpress.WD.ADXWordTaskPanePosition.Right;
            this.adxWordTaskPanesCollectionItem3.RestoreFromMinimizedState = true;
            this.adxWordTaskPanesCollectionItem3.TaskPaneClassName = "ProLoop.Dashboard.ADXWordTaskPaneInfo";
            // 
            // adxWordTaskPanesCollectionItem4
            // 
            this.adxWordTaskPanesCollectionItem4.AlwaysShowHeader = true;
            this.adxWordTaskPanesCollectionItem4.IsHiddenStateAllowed = false;
            this.adxWordTaskPanesCollectionItem4.IsMinimizedStateAllowed = false;
            this.adxWordTaskPanesCollectionItem4.Position = AddinExpress.WD.ADXWordTaskPanePosition.Right;
            this.adxWordTaskPanesCollectionItem4.RestoreFromMinimizedState = true;
            this.adxWordTaskPanesCollectionItem4.TaskPaneClassName = "ProLoop.Dashboard.ADXWordTaskPaneSave";
            // 
            // AddinModule
            // 
            this.AddinName = "ProLoop";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaWord;
            this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);
            this.AddinStartupComplete += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinStartupComplete);

        }
        #endregion

        private AddinExpress.WD.ADXWordTaskPanesManager adxWordTaskPanesManager1;
        private AddinExpress.WD.ADXWordTaskPanesCollectionItem adxWordTaskPanesCollectionItem1;
        private AddinExpress.WD.ADXWordTaskPanesCollectionItem adxWordTaskPanesCollectionItem2;
        private AddinExpress.WD.ADXWordTaskPanesCollectionItem adxWordTaskPanesCollectionItem3;
        private AddinExpress.WD.ADXWordTaskPanesCollectionItem adxWordTaskPanesCollectionItem4;
    }
}

