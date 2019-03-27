[Management.Automation.Runspaces.Runspace]::DefaultRunspace = [RunspaceFactory]::CreateRunspace()

$continuation = [task]::WhenAll($tasks)

$tasks.Add([task]::Factory.StartNew($scriptblock))

[task]::CompletedTask

[task]$tester = New-Object task

[task]::Factory.StartNew | gm *

$test = new-object System.Threading.Tasks.Task($scriptblock)

using namespace System.Collections.Generic
using namespace System.Collections.ObjectModel
using namespace System.Linq.Expressions
using namespace System.Management.Automation
using namespace System.Management.Automation.Runspaces
using namespace System.Threading
using namespace System.Threading.Tasks


$test2 = New-Object Task.factory.StartNew($scriptblock, "beta")