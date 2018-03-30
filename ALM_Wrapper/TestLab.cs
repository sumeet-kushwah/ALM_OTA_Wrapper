
namespace ALM_Wrapper
{
    public class TestLab
    {

        public TestSet TestSet;
        public TestLabFolders TestLabFolders;

        //Set the connection to tests here
        public TestLab(TDAPIOLELib.TDConnection tDConnection)
        {
            TestSet = new TestSet(tDConnection);
            TestLabFolders = new TestLabFolders(tDConnection);
        }

    }
}
