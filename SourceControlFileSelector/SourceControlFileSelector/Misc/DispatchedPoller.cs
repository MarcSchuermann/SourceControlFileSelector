//// --------------------------------------------------------------------------------------------------------------------
//// <copyright>Marc Schürmann</copyright>
//// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Windows.Threading;

namespace SourceControlFileSelector.Misc
{
    public class DispatchedPoller
    {
        #region Public Constructors

        public DispatchedPoller(int maximumNumberOfAttempts, TimeSpan frequency, Func<bool> condition, Action toDo)
        {
            MaximumNumberOfAttempts = maximumNumberOfAttempts;
            Condition = condition;
            ToDo = toDo;
            Frequency = frequency;
        }

        #endregion Public Constructors

        #region Public Properties

        public Func<bool> Condition { get; protected set; }
        public TimeSpan Frequency { get; protected set; }
        public int MaximumNumberOfAttempts { get; protected set; }
        public Action ToDo { get; protected set; }

        #endregion Public Properties

        #region Public Methods

        public void Go()
        {
            Loop();
        }

        #endregion Public Methods

        #region Protected Methods

        protected void Loop()
        {
            if (Condition())
            {
                ToDo();
            }
            else
            {
                int attemptsMade = 0;
                var timer = new DispatcherTimer()
                {
                    Interval = Frequency,
                    Tag = 0
                };
                timer.Tick += (sender, args) =>
                {
                    if (attemptsMade == MaximumNumberOfAttempts)
                    {
                        // Give up, we've tried enough times, no point in continuing
                        timer.Stop();
                    }
                    else
                    {
                        if (Condition())
                        {
                            timer.Stop();
                            ToDo();
                        }
                        else
                        {
                            // Keep the timer going and try again a few more times
                            attemptsMade++;
                        }
                    }
                };
                timer.Start();
            }
        }

        #endregion Protected Methods
    }
}