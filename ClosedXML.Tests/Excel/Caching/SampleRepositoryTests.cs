using ClosedXML.Excel;
using ClosedXML.Excel.Caching;
using NUnit.Framework;
using System.Linq;
using System.Threading.Tasks;

namespace ClosedXML.Tests.Excel.Caching
{
    [TestFixture]
    public class BaseRepositoryTests
    {
        [Test]
        public void DifferentEntitiesWithSameKeyStoredOnce()
        {
            // Arrange
            var key = 12345;
            var entity1 = new SampleEntity(key);
            var entity2 = new SampleEntity(key);
            var sampleRepository = CreateSampleRepository();

            // Act
            var storedEntity1 = sampleRepository.Store(ref key, entity1);
            var storedEntity2 = sampleRepository.Store(ref key, entity2);

            // Assert
            Assert.That(storedEntity1, Is.SameAs(entity1));
            Assert.That(storedEntity2, Is.SameAs(entity1));
            Assert.That(storedEntity2, Is.Not.SameAs(entity2));
        }

        [Test]
        [Explicit("Test reliable fails on build agents since upgraded to .net8, anybody out there who has spare time to fix this? PR welcome")]
        public void NonUsedReferencesAreGCed()
        {
#if !DEBUG
            // Arrange
            int key = 12345;
            var sampleRepository = this.CreateSampleRepository();

            // Act
            var storedEntityRef1 = new System.WeakReference(sampleRepository.Store(ref key, new SampleEntity(key)));

            int count = 0;
            do
            {
                System.Threading.Thread.Sleep(50);
                System.GC.Collect();
                count++;
            } while (storedEntityRef1.IsAlive && count < 10);

            // Assert
           if (count == 10)
                Assert.Fail("storedEntityRef1 was not GCed");

            Assert.IsFalse(sampleRepository.Any());
#else
            Assert.Ignore("Can't run in DEBUG");
#endif
        }

        [Test]
        public void NonUsedReferencesAreGCed2()
        {
#if !DEBUG
            // Arrange
            int countUnique = 30;
            int repeatCount = 1000;
            SampleEntity[] entities = new SampleEntity[countUnique * repeatCount];
            for (int i = 0; i < countUnique; i++)
            {
                for (int j = 0; j < repeatCount; j++)
                {
                    entities[i * repeatCount + j] = new SampleEntity(i);
                }
            }

            var sampleRepository = this.CreateSampleRepository();

            // Act
            Parallel.ForEach(entities, new ParallelOptions { MaxDegreeOfParallelism = 8 },
                e =>
                {
                    var key = e.Key;
                    sampleRepository.Store(ref key, e);
                });

            System.Threading.Thread.Sleep(50);
            System.GC.Collect();
            var storedEntries = sampleRepository.ToList();

            // Assert
            Assert.AreEqual(0, storedEntries.Count);
#else
            Assert.Ignore("Can't run in DEBUG");
#endif
        }

        [Test]
        public void ConcurrentAddingCausesNoDuplication()
        {
            // Arrange
            var countUnique = 30;
            var repeatCount = 1000;
            var entities = new SampleEntity[countUnique * repeatCount];
            for (var i = 0; i < countUnique; i++)
            {
                for (var j = 0; j < repeatCount; j++)
                {
                    entities[i * repeatCount + j] = new SampleEntity(i);
                }
            }

            var sampleRepository = CreateSampleRepository();

            // Act
            Parallel.ForEach(entities, new ParallelOptions { MaxDegreeOfParallelism = 8 },
                e =>
                {
                    var key = e.Key;
                    sampleRepository.Store(ref key, e);
                });
            var storedEntries = sampleRepository.ToList();

            // Assert
            Assert.That(storedEntries.Count, Is.EqualTo(countUnique));
            Assert.That(entities, Is.Not.Null); // To protect them from GC
        }

        [Test]
        public void ReplaceKeyInRepository()
        {
            // Arrange
            var key1 = 12345;
            var key2 = 54321;
            var entity = new SampleEntity(key1);
            var sampleRepository = CreateSampleRepository();
            var storedEntity1 = sampleRepository.Store(ref key1, entity);

            // Act
            sampleRepository.Replace(ref key1, ref key2);
            var containsOld = sampleRepository.ContainsKey(ref key1, out var _);
            var containsNew = sampleRepository.ContainsKey(ref key2, out var _);
            var storedEntity2 = sampleRepository.GetOrCreate(ref key2);

            // Assert
            Assert.That(containsOld, Is.False);
            Assert.That(containsNew, Is.True);
            Assert.That(storedEntity1, Is.SameAs(entity));
            Assert.That(storedEntity2, Is.SameAs(entity));
        }

        [Test]
        public void ConcurrentReplaceKeyInRepository()
        {
            var sampleRepository = new EditableRepository();
            var keys = Enumerable.Range(0, 1000).ToArray();
            keys.ForEach(key => sampleRepository.GetOrCreate(ref key));

            Parallel.ForEach(keys, key =>
            {
                var modifiedKey = key + 2000;
                var val1 = sampleRepository.Replace(ref key, ref modifiedKey);
                val1.Key = key + 2000;
                var val2 = sampleRepository.GetOrCreate(ref modifiedKey);
                Assert.That(val2, Is.SameAs(val1));
            });
        }

        [Test]
        public void ReplaceNonExistingKeyInRepository()
        {
            var key1 = 100;
            var key2 = 200;
            var key3 = 300;
            var entity = new SampleEntity(key1);
            var sampleRepository = CreateSampleRepository();
            sampleRepository.Store(ref key1, entity);

            sampleRepository.Replace(ref key2, ref key3);
            var all = sampleRepository.ToList();

            Assert.That(all, Has.Count.EqualTo(1));
            Assert.That(all.First(), Is.SameAs(entity));
        }

        private SampleRepository CreateSampleRepository()
        {
            return new SampleRepository();
        }

        /// <summary>
        /// Class under testing
        /// </summary>
        internal class SampleRepository : XLRepositoryBase<int, SampleEntity>
        {
            public SampleRepository() : base(key => new SampleEntity(key))
            {
            }
        }

        public class SampleEntity
        {
            public int Key { get; private set; }

            public SampleEntity(int key)
            {
                Key = key;
            }
        }

        /// <summary>
        /// Class under testing
        /// </summary>
        internal class EditableRepository : XLRepositoryBase<int, EditableEntity>
        {
            public EditableRepository() : base(key => new EditableEntity(key))
            {
            }
        }

        public class EditableEntity
        {
            public int Key { get; set; }

            public EditableEntity(int key)
            {
                Key = key;
            }
        }
    }
}