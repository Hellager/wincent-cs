using System;

namespace Wincent
{
    /// <summary>
    /// Controls retry behavior for transient manager shell operations.
    /// </summary>
    public sealed class RetryPolicy
    {
        private static readonly Random JitterRandom = new Random();
        private static readonly object JitterLock = new object();

        /// <summary>
        /// Gets a policy that performs no retries.
        /// </summary>
        public static RetryPolicy None { get; } = new RetryPolicy(0, TimeSpan.Zero, TimeSpan.Zero, 1.0, false);

        /// <summary>
        /// Gets a short retry policy.
        /// </summary>
        public static RetryPolicy Fast { get; } = new RetryPolicy(2, TimeSpan.FromMilliseconds(50), TimeSpan.FromSeconds(1), 1.5, true);

        /// <summary>
        /// Gets the standard retry policy.
        /// </summary>
        public static RetryPolicy Standard { get; } = new RetryPolicy(3, TimeSpan.FromMilliseconds(100), TimeSpan.FromSeconds(5), 2.0, true);

        /// <summary>
        /// Gets an aggressive retry policy.
        /// </summary>
        public static RetryPolicy Aggressive { get; } = new RetryPolicy(5, TimeSpan.FromMilliseconds(200), TimeSpan.FromSeconds(10), 2.0, true);

        /// <summary>
        /// Initializes a retry policy.
        /// </summary>
        public RetryPolicy(int maxRetryCount, TimeSpan initialDelay, TimeSpan maxDelay, double backoffFactor, bool useJitter)
        {
            if (maxRetryCount < 0)
                throw new ArgumentOutOfRangeException(nameof(maxRetryCount));
            if (initialDelay < TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(initialDelay));
            if (maxDelay < TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(maxDelay));
            if (backoffFactor < 1.0)
                throw new ArgumentOutOfRangeException(nameof(backoffFactor));
            if (maxRetryCount > 0 && initialDelay == TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(initialDelay));
            if (maxRetryCount > 0 && maxDelay < initialDelay)
                throw new ArgumentOutOfRangeException(nameof(maxDelay));

            MaxRetryCount = maxRetryCount;
            InitialDelay = initialDelay;
            MaxDelay = maxDelay;
            BackoffFactor = backoffFactor;
            UseJitter = useJitter;
        }

        /// <summary>
        /// Gets the number of retries after the first attempt.
        /// </summary>
        public int MaxRetryCount { get; }

        /// <summary>
        /// Gets the initial retry delay.
        /// </summary>
        public TimeSpan InitialDelay { get; }

        /// <summary>
        /// Gets the maximum retry delay.
        /// </summary>
        public TimeSpan MaxDelay { get; }

        /// <summary>
        /// Gets the exponential backoff factor.
        /// </summary>
        public double BackoffFactor { get; }

        /// <summary>
        /// Gets whether jitter is applied to retry delays.
        /// </summary>
        public bool UseJitter { get; }

        /// <summary>
        /// Gets the delay for a retry attempt.
        /// </summary>
        /// <param name="retryAttempt">The zero-based retry attempt.</param>
        /// <returns>The delay for the retry attempt.</returns>
        /// <exception cref="InvalidOperationException">
        /// This policy does not retry, such as <see cref="None"/>.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <paramref name="retryAttempt"/> is outside the configured retry range.
        /// </exception>
        public TimeSpan GetDelay(int retryAttempt)
        {
            if (MaxRetryCount == 0)
                throw new InvalidOperationException("This retry policy does not retry.");
            if (retryAttempt < 0 || retryAttempt >= MaxRetryCount)
                throw new ArgumentOutOfRangeException(nameof(retryAttempt));

            double milliseconds = InitialDelay.TotalMilliseconds * Math.Pow(BackoffFactor, retryAttempt);
            milliseconds = Math.Min(milliseconds, MaxDelay.TotalMilliseconds);

            if (UseJitter)
            {
                lock (JitterLock)
                {
                    milliseconds *= 0.5 + JitterRandom.NextDouble() * 0.5;
                }
            }

            return TimeSpan.FromMilliseconds(milliseconds);
        }
    }
}
