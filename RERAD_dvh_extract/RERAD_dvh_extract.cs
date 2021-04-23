using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;

namespace RERAD_dvh_extract
{
  class Program
  {
    [STAThread]
    static void Main(string[] args)
    {
      try
      {
        using (Application app = Application.CreateApplication(null, null))
        {
          Execute(app);
        }
      }
      catch (Exception e)
      {
        Console.Error.WriteLine(e.ToString());
      }
    }
        static void Execute(Application app)
        {
            // TODO: add here your code
            // Iterate through all patients
            string outputdir = @"\\Client\E$\rerad\ESAPI_DVH"; // export directory
            //string outputdir = @"\\Client\D$\LeneData\DEPENDS\ESAPI_DVH";
            //string outputdir = @"\\Client\N$\Afdeling\AUHDCENP\MR\Depends\Brain";

            System.IO.Directory.CreateDirectory(outputdir);
            System.IO.File.GetAttributes(outputdir);


            int counter = 0;

            foreach (var patientSummary in app.PatientSummaries)
            {
                // Stop after when a few records have been found
                if (counter > 1000)
                    break;

                // For the test purpose we are interesting in patient records created during the last 6 months
                DateTime startDate = DateTime.Now - TimeSpan.FromDays(30 * 60);
                if (patientSummary.CreationDateTime > startDate)
                {
                    // Retrieve patient information
                    Patient patient = app.OpenPatient(patientSummary);
                    if (patient == null)
                        throw new ApplicationException("Cannot open patient " + patientSummary.Id);

                    String ptID = patient.Id;
                    //if (ptID.Contains("WP12_ID"))

                    //
                    if (patient.LastName.Contains("ReRad")) // the patient lastname must contain ReRad
                    {
                        // Iterate through all patient courses...
                        foreach (var course in patient.Courses)
                        {
                            // ... and plans
                            string message1 = string.Format("Patient: {0}, plans: {1}, course: {2} ", patient.LastName, patient.Id, course.Id);
                            Console.WriteLine(message1);

                            for (int i = 0; i < course.PlanSetups.Count(); i++)
                            {
                                if (course.PlanSums.Count() != 0) //(1 == 0) 
                                {
                                    string message2 = string.Format("Patient: {0}, plans: {1}, has a sumplan", patient.LastName, patient.Id);
                                    Console.WriteLine(message2);
                                    for (int ii = 0; ii < course.PlanSums.Count(); ii++)
                                    {
                                        PlanSum planSum = course.PlanSums.ElementAt(ii);
                                        if ((planSum.Dose != null) && (planSum.Id.ElementAt(0).ToString().Contains("Z")))  // The naming of our sumplan starts with "Z" for sum
                                        {
                                            planSum.DoseValuePresentation = DoseValuePresentation.Absolute;
                                            StructureSet structureSet = planSum.StructureSet;
                                            Structure target = null;
                                            foreach (var structure in structureSet.Structures)
                                            {
                                                if (structure.Id == "Ureter_R")// || structure.Id == "Brainstem")  // here you need to change for every structure
                                                {
                                                    target = structure;
                                                    //int NoParts = target.GetNumberOfSeparateParts();
                                                    break;
                                                }

                                            }

                                            if (target != null)
                                            {    //throw new ApplicationException("The selected plan does not have a Hippocampus_L.");

                                                // Retrieve DVH data
                                                DVHData dvhData = planSum.GetDVHCumulativeData(target, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.001);
                                                if (dvhData != null)
                                                {
                                                    var csv = new StringBuilder();

                                                    foreach (DVHPoint pt in dvhData.CurveData)
                                                    {
                                                        string line = string.Format("{0},{1}", pt.DoseValue.Dose, pt.Volume);
                                                        //Console.WriteLine(line);
                                                        csv.AppendLine(line);
                                                    }
                                                    //string filename = string.Format(@"{0}\{1}\{2}_{3}_{4}_{5}-dvh.csv", outputdir, target.Id, patient.FirstName, course.Id, planSetup.Id, target.Id);
                                                    string filename = string.Format(@"{0}\{1}_Sum\{2}_{3}_{4}_{5}_{6}-dvh.csv", outputdir, target.Id, patient.Id, structureSet.Id, planSum.Id, target.Id, target.GetNumberOfSeparateParts().ToString());
                                                    File.AppendAllText(filename, csv.ToString());

                                                    // ... and display information about max dose
                                                    string message = string.Format("Patient: {0}, Course: {1}, Plan: {2}, Max dose: {3}, target name: {4}, NoParts: {5}", patient.FirstName, course.Id, planSum.Id, planSum.Dose.DoseMax3D.ToString(), target.Id, target.GetNumberOfSeparateParts().ToString());
                                                    Console.WriteLine(message);
                                                }
                                                counter = counter + 1;
                                            }
                                        }

                                    }

                                    break;
                                }
                                PlanSetup planSetup = course.PlanSetups.ElementAt(i);
                                //  try
                                //  {
                                // For the test purpose we will look into approved plans with calculated 3D dose only...
                                PlanSetupApprovalStatus status = planSetup.ApprovalStatus;


                                if ((planSetup.Dose != null) && (status.ToString() != "Rejected"))// && (status.ToString() != "PlanningApproved"))
                                {
                                    if (planSetup.Id.ElementAt(0).ToString().Contains("U"))
                                    {
                                        break;
                                    }
                                    // ... select dose values to be in absolute unit
                                    planSetup.DoseValuePresentation = DoseValuePresentation.Absolute;
                                    StructureSet structureSet = planSetup.StructureSet;
                                    Structure target = null;
                                    foreach (var structure in structureSet.Structures)
                                    {
                                        if (structure.Id == "CTV T")// || structure.Id == "Brainstem")
                                        {
                                            target = structure;
                                            int NoParts = target.GetNumberOfSeparateParts();
                                            break;
                                        }

                                    }

                                    if (target != null)
                                    {    //throw new ApplicationException("The selected plan does not have a Hippocampus_L.");

                                        // Retrieve DVH data
                                        DVHData dvhData = planSetup.GetDVHCumulativeData(target, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.001);
                                        if (dvhData != null)
                                        {
                                            var csv = new StringBuilder();

                                            foreach (DVHPoint pt in dvhData.CurveData)
                                            {
                                                string line = string.Format("{0},{1}", pt.DoseValue.Dose, pt.Volume);
                                                //Console.WriteLine(line);
                                                csv.AppendLine(line);
                                            }
                                            //string filename = string.Format(@"{0}\{1}\{2}_{3}_{4}_{5}-dvh.csv", outputdir, target.Id, patient.FirstName, course.Id, planSetup.Id, target.Id);
                                            string filename = string.Format(@"{0}\{1}\{2}_{3}_{4}_{5}_{6}-dvh.csv", outputdir, target.Id, patient.Id, structureSet.Id, planSetup.Id, target.Id, target.GetNumberOfSeparateParts().ToString());
                                            File.AppendAllText(filename, csv.ToString());
                                            
                                            // ... and display information about max dose
                                            string message = string.Format("Patient: {0}, Course: {1}, Plan: {2}, Max dose: {3}, target name: {4}, NoParts: {5}", patient.FirstName, course.Id, planSetup.Id, planSetup.Dose.DoseMax3D.ToString(), target.Id, target.GetNumberOfSeparateParts().ToString());
                                            Console.WriteLine(message);
                                        }
                                        counter = counter + 1;
                                    }
                                }
                            }
                        }
                    }
                    app.ClosePatient();
                }
            }
        }
    }
}
