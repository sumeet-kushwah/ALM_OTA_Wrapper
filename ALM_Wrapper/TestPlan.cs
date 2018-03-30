
namespace ALM_Wrapper
{
    public class TestPlan 
    {
        public Test Test;
        public TestFolders TestFolders;

        //Set the connection to tests here
        public TestPlan(TDAPIOLELib.TDConnection OALMConnection)
        {
            Test = new Test(OALMConnection);
            TestFolders = new TestFolders(OALMConnection);
        }
    }
}
