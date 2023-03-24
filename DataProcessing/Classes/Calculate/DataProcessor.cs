using DataProcessing.Classes.Calculate;
using DataProcessing.Constants;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace DataProcessing.Classes.Calculate
{
    enum GraphTableDataType
    {
        Seconds,
        Minutes,
        Percentages,
        Numbers
    }

    internal class DataProcessor
    {
        #region Private attributes
        private readonly CalculationOptions options;
        private readonly CalculatedData calculatedData;
        private readonly Calculator calculator;
        #endregion

        #region Constructors
        public DataProcessor(CalculationOptions options)
        {
            // Init
            this.options = options;
            calculatedData = new CalculatedData();            
            calculator = new Calculator();

            // Extract all distinct states from excel file
            List<int> states = options.MarkedTimeStamps
                                            .Where(sample => sample.State != 0)
                                            .Select(sample => sample.State)
                                            .Distinct()
                                            .ToList();
            states.Sort();

            // If number of extracted states doesn't match number of selected states throw error.
            if (states.Count > options.MaxStates) 
            { 
                throw new Exception($"File contains more than {options.MaxStates} states!"); 
            }

            calculatedData.CreatePhases(options.MaxStates);
        }
        #endregion

        #region Public methods
        public CalculatedData Calculate()
        {
            // Create duplicated timestamps for graph
            calculatedData.duplicatedTimes = calculator.generateDuplicatedTimeStamps(options.MarkedTimeStamps);

            // Calculate total
            calculatedData.totalStats = calculator.CalculateStats(options.NonMarkedTimeStamps, calculatedData.GetStates(), options.Criterias);


            // Add total here so it will be on top of hourly frequencies
            calculatedData.AddFrequency(calculator.calculateFrequencies(options.NonMarkedTimeStamps, calculatedData.GetStates()));
            calculatedData.AddFrequencyRange(calculator.calculateFrequencyRanges(options.NonMarkedTimeStamps, calculatedData.GetStates(), options.FrequencyRanges));

            // Latency
            calculatedData.timeBeforeFirstSleep = calculator.calculateStateLatency(options.MarkedTimeStamps, options.GetState("Sleep"));
            if (options.MaxStates == 3)
            {
                calculatedData.timeBeforeFirstParadoxicalSleep = calculator.calculateStateLatency(options.MarkedTimeStamps, options.GetState("Paradoxical sleep"));
            }

            // Calculate per hour
            int time = 0;
            int currentHour = 0;
            List<TimeStamp> hourRegion = new List<TimeStamp>();
            for (int i = 0; i < options.MarkedTimeStamps.Count; i++)
            {
                TimeStamp currentTimeStamp = options.MarkedTimeStamps[i];
                time += currentTimeStamp.TimeDifferenceInSeconds;

                if (time > options.TimeMarkInSeconds) { throw new Exception("Invalid hour marks"); }

                hourRegion.Add(currentTimeStamp);

                if (time == options.TimeMarkInSeconds)
                {
                    currentHour++;
                    calculatedData.hourAndStats.Add(currentHour, calculator.CalculateStats(hourRegion, options.GetAllStates(), options.Criterias));
                    calculatedData.AddFrequency(calculator.calculateFrequencies(hourRegion, options.GetAllStates()));
                    calculatedData.AddFrequencyRange(calculator.calculateFrequencyRanges(hourRegion, options.GetAllStates(), options.FrequencyRanges));

                    time = 0;
                    hourRegion.Clear();
                }
            }

            // Do last part (might be less than marked time)
            if (hourRegion.Count != 0)
            {
                currentHour++;
                calculatedData.hourAndStats.Add(currentHour, calculator.CalculateStats(hourRegion, options.GetAllStates(), options.Criterias));
                calculatedData.AddFrequency(calculator.calculateFrequencies(hourRegion, options.GetAllStates()));
                calculatedData.AddFrequencyRange(calculator.calculateFrequencyRanges(hourRegion, options.GetAllStates(), options.FrequencyRanges));
            }

            // Calculate stats for clusters
            if (options.ClusterSeparationTimeInSeconds > 0)
            {
                int curClusterNumber = 1;
                foreach (List<TimeStamp> cluster in calculator.CreateClusters(options.NonMarkedTimeStamps, options.ClusterSeparationTimeInSeconds, options.GetState("Wakefulness")))
                {
                    calculatedData.clusterAndStats.Add(curClusterNumber, calculator.CalculateStats(cluster, options.GetAllStates(), options.Criterias));
                    curClusterNumber++;
                }
            }

            return calculatedData;
        }
        #endregion
    }
}
