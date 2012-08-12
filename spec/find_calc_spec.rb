require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe "IAccessibleClient" do
  before(:each) do
    @sample_app = 'calc.exe'
    @pid1 = IO.popen(@sample_app).pid
    sleep 1
  end

  after(:each) do
    Process.kill(9, @pid1) rescue nil
  end

  it "finds calc.exe element" do
    desktop = IAccessibleClient.desktop
    desktop.name.should == 'Desktop'
    calc = desktop.find_child('Calculator')
    calc.name.should == 'Calculator'
    # calc.iterate_accessible_children_recursively
  end
end

