using System;
using System.Collections.Generic;
using Microsoft.Build.Framework;

namespace DotNetCoreTypeLibGenerator.Tests
{
    public class FakeBuildEngine : IBuildEngine
    {
        // It's just a test helper so public fields is fine.
        public List<BuildErrorEventArgs> LogErrorEvents = new List<BuildErrorEventArgs>();

        public List<BuildMessageEventArgs> LogMessageEvents =
            new List<BuildMessageEventArgs>();

        public List<CustomBuildEventArgs> LogCustomEvents =
            new List<CustomBuildEventArgs>();

        public List<BuildWarningEventArgs> LogWarningEvents =
            new List<BuildWarningEventArgs>();

        public virtual bool BuildProjectFile(
            string projectFileName, string[] targetNames,
            System.Collections.IDictionary globalProperties,
            System.Collections.IDictionary targetOutputs)
        {
            throw new NotImplementedException();
        }

        public virtual int ColumnNumberOfTaskNode
        {
            get { return 0; }
        }

        public virtual bool ContinueOnError
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public virtual int LineNumberOfTaskNode
        {
            get { return 0; }
        }

        public virtual void LogCustomEvent(CustomBuildEventArgs e)
        {
            LogCustomEvents.Add(e);
        }

        public virtual void LogErrorEvent(BuildErrorEventArgs e)
        {
            LogErrorEvents.Add(e);
        }

        public virtual void LogMessageEvent(BuildMessageEventArgs e)
        {
            LogMessageEvents.Add(e);
        }

        public virtual void LogWarningEvent(BuildWarningEventArgs e)
        {
            LogWarningEvents.Add(e);
        }

        public virtual string ProjectFileOfTaskNode
        {
            get { return "fake ProjectFileOfTaskNode"; }
        }

    }
}
