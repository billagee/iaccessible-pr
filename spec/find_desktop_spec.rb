require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

# Demo script that asserts the name of the root (desktop) UI element using MSAA

describe "IAccessibleClient" do
  it "finds desktop element" do
    desktop = IAccessibleClient.desktop
    desktop.name.should == 'Desktop'
  end
end

