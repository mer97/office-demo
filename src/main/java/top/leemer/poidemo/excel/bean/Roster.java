package top.leemer.poidemo.excel.bean;

/**
 * @author LEEMER
 * Create Date: 2019-10-09
 */
public class Roster {

    private String id;

    private String name;

    private String phone;

    private String site;

    public Roster() {
    }

    public Roster(String id, String name, String phone, String site) {
        this.id = id;
        this.name = name;
        this.phone = phone;
        this.site = site;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public String getSite() {
        return site;
    }

    public void setSite(String site) {
        this.site = site;
    }

    @Override
    public String toString() {
        return "Roster{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", phone='" + phone + '\'' +
                ", site='" + site + '\'' +
                '}';
    }
}
