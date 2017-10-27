package me.qinmian.test.bean;

import me.qinmian.annotation.IgnoreField;

public class Role {
	
	private String roleName;
	
	private String roleDesc;

	@IgnoreField
	private Privilege privilege;
	
	public Role() {
		super();
	}

	public Role(String roleName,String roleDesc) {
		super();
		this.roleName = roleName;
		this.roleDesc = roleDesc;
	}

	public String getRoleName() {
		return roleName;
	}

	public void setRoleName(String roleName) {
		this.roleName = roleName;
	}

	public String getRoleDesc() {
		return roleDesc;
	}

	public void setRoleDesc(String roleDesc) {
		this.roleDesc = roleDesc;
	}

	public Privilege getPrivilege() {
		return privilege;
	}

	public void setPrivilege(Privilege privilege) {
		this.privilege = privilege;
	}
	
}
